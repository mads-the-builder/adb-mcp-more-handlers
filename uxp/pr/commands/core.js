const app = require("premierepro");
const constants = require("premierepro").Constants;

const {BLEND_MODES, TRACK_TYPE } = require("./consts.js")

const {
    _getSequenceFromId,
    _setActiveSequence,
    _openSequence,
    setParam,
    getParam,
    addEffect,
    findProjectItem,
    findProjectItemAtRoot,
    findProjectItemByPath,
    findBinByPath,
    namesMatchWithVersionNormalization,
    execute,
    getTrack,
    getTrackItems
} = require("./utils.js")

const saveProject = async (command) => {
    let project = await app.Project.getActiveProject()

    project.save()
}

const saveProjectAs = async (command) => {
    let project = await app.Project.getActiveProject()

    const options = command.options;
    const filePath = options.filePath;

    project.saveAs(filePath)
}

const openProject = async (command) => {

    const options = command.options;
    const filePath = options.filePath;

    await app.Project.open(filePath);    
}


const importMedia = async (command) => {

    let options = command.options
    let paths = command.options.filePaths

    let project = await app.Project.getActiveProject()

    let root = await project.getRootItem()
    let originalItems = await root.getItems()

    let success = await project.importFiles(paths, true, root)

    let updatedItems = await root.getItems()
    
    const addedItems = updatedItems.filter(
        updatedItem => !originalItems.some(originalItem => originalItem.name === updatedItem.name)
      );
      
    let addedProjectItems = [];
    for (const p of addedItems) { 
        addedProjectItems.push({ name: p.name });
    }
    
    return { addedProjectItems };
}


const addMediaToSequence = async (command) => {

    let options = command.options
    let itemName = options.itemName
    let id = options.sequenceId
    const videoOnly = options.audioTrackIndex === -1

    let project = await app.Project.getActiveProject()
    let sequence = await _getSequenceFromId(id)

    let insertItem = await findProjectItem(itemName, project)

    let editor = await app.SequenceEditor.getEditor(sequence)

    const insertionTime = await app.TickTime.createWithTicks(options.insertionTimeTicks.toString());
    const videoTrackIndex = options.videoTrackIndex
    let audioTrackIndex = options.audioTrackIndex

    // VIDEO-ONLY MODE: Find a safe audio track, insert, then remove audio
    if (videoOnly) {
        // Find an available audio track (one with no clips in our time range)
        const startTicks = BigInt(options.insertionTimeTicks);
        // Estimate end time - we'll use a large range to be safe
        // The clip will be pre-trimmed so we use that duration
        const endTicks = startTicks + BigInt(254016000000 * 60); // ~60 seconds buffer

        const audioTrackCount = await sequence.getAudioTrackCount();
        let safeAudioTrack = -1;

        // Search for an empty audio track (start from highest to avoid common tracks)
        for (let i = Math.max(0, audioTrackCount - 1); i >= 0; i--) {
            const track = await sequence.getAudioTrack(i);
            if (!track) continue;

            const clips = await track.getTrackItems(1, false);
            if (!clips || clips.length === 0) {
                safeAudioTrack = i;
                break;
            }

            // Check if track is empty in our time range
            let hasOverlap = false;
            for (const clip of clips) {
                const clipStartTime = await clip.getStartTime();
                const clipEndTime = await clip.getEndTime();
                const clipStart = BigInt(clipStartTime.ticksNumber);
                const clipEnd = BigInt(clipEndTime.ticksNumber);

                if (clipStart < endTicks && clipEnd > startTicks) {
                    hasOverlap = true;
                    break;
                }
            }

            if (!hasOverlap) {
                safeAudioTrack = i;
                break;
            }
        }

        // If no safe track found, use track index = audioTrackCount (will create new track)
        if (safeAudioTrack === -1) {
            safeAudioTrack = audioTrackCount;
        }

        audioTrackIndex = safeAudioTrack;
    }

    // Insert the clip
    await execute(() => {
        let action = editor.createOverwriteItemAction(insertItem, insertionTime, videoTrackIndex, audioTrackIndex)
        return [action]
    }, project)

    // VIDEO-ONLY MODE: Remove the audio clip we just inserted
    if (videoOnly) {
        // Find and remove the audio clip at the insertion time on the safe audio track
        const audioTrack = await sequence.getAudioTrack(audioTrackIndex);
        if (audioTrack) {
            const clips = await audioTrack.getTrackItems(1, false);
            const insertionTicks = BigInt(options.insertionTimeTicks);

            for (const clip of clips) {
                const clipStartTime = await clip.getStartTime();
                const clipStart = BigInt(clipStartTime.ticksNumber);

                // Find clip that starts at or very close to our insertion time
                if (clipStart >= insertionTicks - BigInt(1000) && clipStart <= insertionTicks + BigInt(1000)) {
                    // Get selection from sequence and clear it, then add our clip
                    const selection = await sequence.getSelection();
                    const existingItems = await selection.getTrackItems();
                    for (const item of existingItems) {
                        await selection.removeItem(item);
                    }
                    await selection.addItem(clip, true);

                    await execute(() => {
                        const removeAction = editor.createRemoveItemsAction(selection, false, constants.MediaType.AUDIO, false);
                        return [removeAction];
                    }, project);

                    break;
                }
            }
        }
    }
}


const setAudioTrackMute = async (command) => {

    let options = command.options
    let id = options.sequenceId

    let sequence = await _getSequenceFromId(id)

    let track = await sequence.getTrack(options.audioTrackIndex, TRACK_TYPE.AUDIO)
    track.setMute(options.mute)
}



const setVideoClipProperties = async (command) => {

    const options = command.options
    let id = options.sequenceId

    let project = await app.Project.getActiveProject()
    let sequence = await _getSequenceFromId(id)

    if(!sequence) {
        throw new Error(`setVideoClipProperties : Requires an active sequence.`)
    }

    let trackItem = await getTrack(sequence, options.videoTrackIndex, options.trackItemIndex, TRACK_TYPE.VIDEO)

    let opacityParam = await getParam(trackItem, "AE.ADBE Opacity", "Opacity")
    let opacityKeyframe = await opacityParam.createKeyframe(options.opacity)

    let blendModeParam = await getParam(trackItem, "AE.ADBE Opacity", "Blend Mode")

    let mode = BLEND_MODES[options.blendMode.toUpperCase()]
    let blendModeKeyframe = await blendModeParam.createKeyframe(mode)

    execute(() => {
        let opacityAction = opacityParam.createSetValueAction(opacityKeyframe);
        let blendModeAction = blendModeParam.createSetValueAction(blendModeKeyframe);
        return [opacityAction, blendModeAction]
    }, project)
}

const appendVideoFilter = async (command) => {

    let options = command.options
    let id = options.sequenceId

    let sequence = await _getSequenceFromId(id)

    if(!sequence) {
        throw new Error(`appendVideoFilter : Requires an active sequence.`)
    }

    let trackItem = await getTrack(sequence, options.videoTrackIndex, options.trackItemIndex, TRACK_TYPE.VIDEO)

    let effectName = options.effectName
    let properties = options.properties

    await addEffect(trackItem, effectName)

    for(const p of properties) {
        await setParam(trackItem, effectName, p.name, p.value)
    }
}


const setActiveSequence = async (command) => {
    let options = command.options
    let id = options.sequenceId

    let sequence = await _getSequenceFromId(id)

    await _setActiveSequence(sequence)
}

const openSequence = async (command) => {
    let options = command.options
    let id = options.sequenceId

    let sequence = await _getSequenceFromId(id)

    await _openSequence(sequence)
}

const createProject = async (command) => {

    let options = command.options
    let path = options.path
    let name = options.name

    if (!path.endsWith('/')) {
        path = path + '/';
    }

    let project = await app.Project.createProject(`${path}${name}.prproj`)

    if(!project) {
        throw new Error("createProject : Could not create project. Check that the directory path exists and try again.")
    }
}


const _exportFrame = async (sequence, filePath, seconds) => {

    const fileType = filePath.split('.').pop()

    let size = await sequence.getFrameSize()

    let p = window.path.parse(filePath)
    let t = app.TickTime.createWithSeconds(seconds)

    let out = await app.Exporter.exportSequenceFrame(sequence, t, p.base, p.dir, size.width, size.height)

    let ps = `${p.dir}${window.path.sep}${p.base}`
    let outPath = `${ps}.${fileType}`

    if(!out) {
        throw new Error(`exportFrame : Could not save frame to [${outPath}]`);
    }

    return outPath
}

const exportFrame = async (command) => {
    const options = command.options;
    let id = options.sequenceId;
    let filePath = options.filePath;
    let seconds = options.seconds;

    let sequence = await _getSequenceFromId(id);

    const outPath = await _exportFrame(sequence, filePath, seconds);

    return {"filePath": outPath}
}

const setClipDisabled = async (command) => {

    const options = command.options;
    const id = options.sequenceId;
    const trackIndex = options.trackIndex;
    const trackItemIndex = options.trackItemIndex;
    const trackType = options.trackType;

    let project = await app.Project.getActiveProject()
    let sequence = await _getSequenceFromId(id)

    if(!sequence) {
        throw new Error(`setClipDisabled : Requires an active sequence.`)
    }

    let trackItem = await getTrack(sequence, trackIndex, trackItemIndex, trackType)

    execute(() => {
        let action = trackItem.createSetDisabledAction(options.disabled)
        return [action]
    }, project)

}


const appendVideoTransition = async (command) => {

    let options = command.options
    let id = options.sequenceId

    let project = await app.Project.getActiveProject()
    let sequence = await _getSequenceFromId(id)

    if(!sequence) {
        throw new Error(`appendVideoTransition : Requires an active sequence.`)
    }

    let trackItem = await getTrack(sequence, options.videoTrackIndex, options.trackItemIndex,TRACK_TYPE.VIDEO)

    let transition = await app.TransitionFactory.createVideoTransition(options.transitionName);

    let transitionOptions = new app.AddTransitionOptions()
    transitionOptions.setApplyToStart(false)

    const time = await app.TickTime.createWithSeconds(options.duration)
    transitionOptions.setDuration(time)
    transitionOptions.setTransitionAlignment(options.clipAlignment)

    execute(() => {
        let action = trackItem.createAddVideoTransitionAction(transition, transitionOptions)
        return [action]
    }, project)
}


const getProjectInfo = async (command) => {
    return {}
}



const createSequenceFromMedia = async (command) => {

    let options = command.options

    let itemNames = options.itemNames
    let sequenceName = options.sequenceName

    let project = await app.Project.getActiveProject()

    let found = false
    try {
        await findProjectItem(sequenceName, project)
        found  = true
    } catch {
    }

    if(found) {
        throw Error(`createSequenceFromMedia : sequence name [${sequenceName}] is already in use`)
    }

    let items = []
    for (const name of itemNames) {
        let insertItem = await findProjectItem(name, project)
        items.push(insertItem)
    }


    let root = await project.getRootItem()
    
    let sequence = await project.createSequenceFromMedia(sequenceName, items, root)

    await _setActiveSequence(sequence)
}

const setClipStartEndTimes = async (command) => {
    const options = command.options;

    const sequenceId = options.sequenceId;
    const trackIndex = options.trackIndex;
    const trackItemIndex = options.trackItemIndex;
    const startTimeTicks = options.startTimeTicks;
    const endTimeTicks = options.endTimeTicks;
    const trackType = options.trackType

    const sequence = await _getSequenceFromId(sequenceId)
    let trackItem = await getTrack(sequence, trackIndex, trackItemIndex, trackType)

    const startTick = await app.TickTime.createWithTicks(startTimeTicks.toString());
    const endTick = await app.TickTime.createWithTicks(endTimeTicks.toString());;

    let project = await app.Project.getActiveProject();

    execute(() => {

        let out = []

        out.push(trackItem.createSetStartAction(startTick));
        out.push(trackItem.createSetEndAction(endTick))

        return out
    }, project)
}

const setClipInOutPoints = async (command) => {
    const options = command.options;

    const sequenceId = options.sequenceId;
    const trackIndex = options.trackIndex;
    const trackItemIndex = options.trackItemIndex;
    const inPointTicks = options.inPointTicks;
    const outPointTicks = options.outPointTicks;
    const trackType = options.trackType;

    const sequence = await _getSequenceFromId(sequenceId);
    let trackItem = await getTrack(sequence, trackIndex, trackItemIndex, trackType);

    const inTick = await app.TickTime.createWithTicks(inPointTicks.toString());
    const outTick = await app.TickTime.createWithTicks(outPointTicks.toString());

    let project = await app.Project.getActiveProject();

    execute(() => {
        let out = [];
        out.push(trackItem.createSetInPointAction(inTick));
        out.push(trackItem.createSetOutPointAction(outTick));
        return out;
    }, project);
}

const moveClip = async (command) => {
    const options = command.options;

    const sequenceId = options.sequenceId;
    const trackIndex = options.trackIndex;
    const trackItemIndex = options.trackItemIndex;
    const moveByTicks = options.moveByTicks;
    const trackType = options.trackType;

    const sequence = await _getSequenceFromId(sequenceId);
    let trackItem = await getTrack(sequence, trackIndex, trackItemIndex, trackType);

    const shiftTick = await app.TickTime.createWithTicks(moveByTicks.toString());

    let project = await app.Project.getActiveProject();

    execute(() => {
        let out = [];
        out.push(trackItem.createMoveAction(shiftTick));
        return out;
    }, project);
}

const closeGapsOnSequence = async(command) => {
    const options = command.options
    const sequenceId = options.sequenceId;
    const trackIndex = options.trackIndex;
    const trackType = options.trackType;

    let sequence = await _getSequenceFromId(sequenceId)

    let out = await _closeGapsOnSequence(sequence, trackIndex, trackType)
    
    return out
}

const _closeGapsOnSequence = async (sequence, trackIndex, trackType) => {
  
    let project = await app.Project.getActiveProject()

    let items = await getTrackItems(sequence, trackIndex, trackType)

    if(!items || items.length === 0) {
        return;
    }
    
    const f = async (item, targetPosition) => {
        let currentStart = await item.getStartTime()

        let a = await currentStart.ticksNumber
        let b = await targetPosition.ticksNumber
        let shiftAmount = (a - b)// How much to shift 
        
        shiftAmount *= -1;

        let shiftTick = app.TickTime.createWithTicks(shiftAmount.toString())

        return shiftTick
    }

    let targetPosition = app.TickTime.createWithTicks("0")


    for(let i = 0; i < items.length; i++) {
        let item = items[i];
        let shiftTick = await f(item, targetPosition)
        
        execute(() => {
            let out = []

                out.push(item.createMoveAction(shiftTick))

            return out
        }, project)
        
        targetPosition = await item.getEndTime()
    }
}

const removeItemFromSequence = async (command) => {
    const options = command.options;

    const sequenceId = options.sequenceId;
    const trackIndex = options.trackIndex;
    const trackItemIndex = options.trackItemIndex;
    const rippleDelete = options.rippleDelete;
    const trackType = options.trackType

    let project = await app.Project.getActiveProject()
    let sequence = await _getSequenceFromId(sequenceId)

    if(!sequence) {
        throw Error(`removeItemFromSequence : sequence with id [${sequenceId}] not found.`)
    }

    let item = await getTrack(sequence, trackIndex, trackItemIndex, trackType);

    let editor = await app.SequenceEditor.getEditor(sequence)

    let trackItemSelection = await sequence.getSelection();
    let items = await trackItemSelection.getTrackItems()

    for (let t of items) {
        await trackItemSelection.removeItem(t)
    }

    trackItemSelection.addItem(item, true)

    execute(() => {
        const shiftOverlapping = false
        let action = editor.createRemoveItemsAction(trackItemSelection, rippleDelete, constants.MediaType.ANY, shiftOverlapping )
        return [action]
    }, project)
}

const addMarkerToSequence = async (command) => {
    const options = command.options;
    const sequenceId = options.sequenceId;
    const markerName = options.markerName;
    const startTimeTicks = options.startTimeTicks;
    const durationTicks = options.durationTicks;
    const comments = options.comments;

    const sequence = await _getSequenceFromId(sequenceId)

    if(!sequence) {
        throw Error(`addMarkerToSequence : sequence with id [${sequenceId}] not found.`)
    }

    let markers = await app.Markers.getMarkers(sequence);

    let project = await app.Project.getActiveProject()

    execute(() => {

        let start = app.TickTime.createWithTicks(startTimeTicks.toString())
        let duration = app.TickTime.createWithTicks(durationTicks.toString())

        let action = markers.createAddMarkerAction(markerName, "WebLink",  start, duration, comments)
        return [action]
    }, project)

}

const moveProjectItemsToBin = async (command) => {
    const options = command.options;
    const binName = options.binName;
    const projectItemNames = options.itemNames;

    const project = await app.Project.getActiveProject()

    // Use path-based lookup if binName contains '/', otherwise simple name lookup
    let binFolderItem;
    if (binName.includes('/')) {
        binFolderItem = await findBinByPath(binName, project);
    } else {
        binFolderItem = await findProjectItem(binName, project);
    }

    if(!binFolderItem) {
        throw Error(`moveProjectItemsToBin : Bin with name [${binName}] not found.`)
    }

    // Get root items ONCE for fast local lookup
    const rootFolderItem = await project.getRootItem();
    const rootItems = await rootFolderItem.getItems();

    let folderItems = [];

    for(let name of projectItemNames) {
        // Fast path: search locally in already-fetched root items
        let item = rootItems.find(i =>
            namesMatchWithVersionNormalization(name, i.name)
        );

        // Slow path: full recursive search if not found at root
        if (!item) {
            item = await findProjectItem(name, project);
        }

        if(!item) {
            throw Error(`moveProjectItemsToBin : FolderItem with name [${name}] not found.`)
        }

        folderItems.push(item)
    }

    execute(() => {

        let actions = []

        for(let folderItem of folderItems) {
            let b = app.FolderItem.cast(binFolderItem)
            let action = rootFolderItem.createMoveItemAction(folderItem, b)
            actions.push(action)
        }

        return actions
    }, project)

}

/**
 * Batch move multiple items to multiple different bins in ONE call.
 * options.moves = [{itemName: "file1", binPath: "vfx/shot1"}, ...]
 */
const batchMoveItemsToBins = async (command) => {
    const options = command.options;
    const moves = options.moves; // [{itemName, binPath}, ...]

    const project = await app.Project.getActiveProject();
    const rootFolderItem = await project.getRootItem();
    const rootItems = await rootFolderItem.getItems();

    // Collect all unique bin paths and resolve them
    const binPaths = [...new Set(moves.map(m => m.binPath))];
    const binCache = new Map();

    for (const binPath of binPaths) {
        let bin;
        if (binPath.includes('/')) {
            bin = await findBinByPath(binPath, project);
        } else {
            bin = await findProjectItem(binPath, project);
        }
        if (bin) {
            binCache.set(binPath, bin);
        }
    }

    // Resolve all items and pair with their target bins
    const moveActions = [];

    for (const move of moves) {
        const { itemName, binPath } = move;

        // Find item (fast path: root level)
        let item = rootItems.find(i =>
            namesMatchWithVersionNormalization(itemName, i.name)
        );
        if (!item) {
            item = await findProjectItem(itemName, project);
        }

        const targetBin = binCache.get(binPath);

        if (item && targetBin) {
            moveActions.push({ item, targetBin });
        }
    }

    // Execute all moves in one batch
    execute(() => {
        let actions = [];
        for (const { item, targetBin } of moveActions) {
            let b = app.FolderItem.cast(targetBin);
            let action = rootFolderItem.createMoveItemAction(item, b);
            actions.push(action);
        }
        return actions;
    }, project);
}

const createBinInActiveProject = async (command) => {
    const options = command.options;
    const binName = options.binName;

    const project = await app.Project.getActiveProject()
    const folderItem = await project.getRootItem()

    // Check if bin already exists to avoid duplicates
    const existingItems = await folderItem.getItems();
    const alreadyExists = existingItems.some(item =>
        namesMatchWithVersionNormalization(binName, item.name) &&
        app.FolderItem.cast(item)
    );

    if (alreadyExists) {
        return; // Bin already exists, don't create duplicate
    }

    execute(() => {
        let action = folderItem.createBinAction(binName, true)
        return [action]
    }, project)
}

const createBinInBin = async (command) => {
    const options = command.options;
    const parentBinName = options.parentBinName;
    const binName = options.binName;

    const project = await app.Project.getActiveProject()

    // Find the parent bin by name
    let parentBin;
    try {
        parentBin = await findProjectItem(parentBinName, project)
    } catch {
        throw Error(`createBinInBin: parent bin "${parentBinName}" not found`)
    }

    if (!parentBin) {
        throw Error(`createBinInBin: parent bin "${parentBinName}" not found`)
    }

    // Verify it's a folder/bin
    const parentFolder = app.FolderItem.cast(parentBin)
    if (!parentFolder) {
        throw Error(`createBinInBin: "${parentBinName}" is not a bin/folder`)
    }

    // Check if bin already exists to avoid duplicates
    const existingItems = await parentFolder.getItems();
    const alreadyExists = existingItems.some(item =>
        namesMatchWithVersionNormalization(binName, item.name) &&
        app.FolderItem.cast(item)
    );

    if (alreadyExists) {
        return; // Bin already exists, don't create duplicate
    }

    execute(() => {
        let action = parentFolder.createBinAction(binName, true)
        return [action]
    }, project)
}

const exportSequence = async (command) => {
    const options = command.options;
    const sequenceId = options.sequenceId;
    const outputPath = options.outputPath;
    const presetPath = options.presetPath;

    const manager = await app.EncoderManager.getManager();

    const sequence = await _getSequenceFromId(sequenceId);

    await manager.exportSequence(sequence, constants.ExportType.IMMEDIATELY, outputPath, presetPath);
}

const { sendEvent } = require("./events.js");

/**
 * Export sequence with progress events.
 *
 * Same as exportSequence but sends progress/completion events via the event channel.
 * Handler still returns normally when complete (same pattern as other handlers).
 */
const exportSequenceWithProgress = async (command) => {
    const options = command.options;
    const sequenceId = options.sequenceId;
    const outputPath = options.outputPath;
    const presetPath = options.presetPath;
    const jobId = options.jobId || `export_${Date.now()}`;

    const manager = await app.EncoderManager.getManager();
    const sequence = await _getSequenceFromId(sequenceId);

    // Track completion state
    let completed = false;
    let exportError = null;

    // Event handlers
    const onProgress = (event) => {
        sendEvent('export_progress', {
            jobId,
            progress: event?.progress,
            outputPath
        });
    };

    const onComplete = (event) => {
        completed = true;
        sendEvent('export_complete', {
            jobId,
            outputPath,
            success: true
        });
    };

    const onError = (event) => {
        completed = true;
        exportError = event?.message || 'Unknown export error';
        sendEvent('export_error', {
            jobId,
            outputPath,
            error: exportError
        });
    };

    const onCancel = (event) => {
        completed = true;
        exportError = 'Export cancelled';
        sendEvent('export_cancel', {
            jobId,
            outputPath
        });
    };

    // Register event listeners
    try {
        app.addEventListener(manager, constants.EncoderManager.EVENT_RENDER_PROGRESS, onProgress);
        app.addEventListener(manager, constants.EncoderManager.EVENT_RENDER_COMPLETE, onComplete);
        app.addEventListener(manager, constants.EncoderManager.EVENT_RENDER_ERROR, onError);
        app.addEventListener(manager, constants.EncoderManager.EVENT_RENDER_CANCEL, onCancel);
    } catch (e) {
        console.warn('[exportSequenceWithProgress] Could not register event listeners:', e);
        // Fall back to non-event export
        await manager.exportSequence(sequence, constants.ExportType.IMMEDIATELY, outputPath, presetPath);
        return { jobId, outputPath, success: true, eventsSupported: false };
    }

    // Start export
    try {
        await manager.exportSequence(sequence, constants.ExportType.IMMEDIATELY, outputPath, presetPath);
    } finally {
        // Clean up event listeners
        try {
            app.removeEventListener(manager, constants.EncoderManager.EVENT_RENDER_PROGRESS, onProgress);
            app.removeEventListener(manager, constants.EncoderManager.EVENT_RENDER_COMPLETE, onComplete);
            app.removeEventListener(manager, constants.EncoderManager.EVENT_RENDER_ERROR, onError);
            app.removeEventListener(manager, constants.EncoderManager.EVENT_RENDER_CANCEL, onCancel);
        } catch (e) {
            console.warn('[exportSequenceWithProgress] Could not remove event listeners:', e);
        }
    }

    if (exportError) {
        throw new Error(exportError);
    }

    return { jobId, outputPath, success: true, eventsSupported: true };
}

const cloneSequence = async (command) => {
    const options = command.options;
    const sequenceId = options.sequenceId;

    const project = await app.Project.getActiveProject();
    const sequence = await _getSequenceFromId(sequenceId);

    // Get sequences before cloning to identify the new one
    const sequencesBefore = await project.getSequences();
    const guidsBefore = new Set(sequencesBefore.map(s => s.guid.toString()));

    // Execute clone action
    execute(() => {
        const action = sequence.createCloneAction();
        return [action];
    }, project);

    // Find the newly created sequence
    const sequencesAfter = await project.getSequences();
    let clonedSequence = null;
    for (const seq of sequencesAfter) {
        if (!guidsBefore.has(seq.guid.toString())) {
            clonedSequence = seq;
            break;
        }
    }

    if (!clonedSequence) {
        throw new Error("cloneSequence: Could not find cloned sequence after operation");
    }

    return {
        id: clonedSequence.guid.toString(),
        name: clonedSequence.name
    };
}

const renameSequence = async (command) => {
    const options = command.options;
    const sequenceId = options.sequenceId;
    const newName = options.newName;

    const project = await app.Project.getActiveProject();
    const sequence = await _getSequenceFromId(sequenceId);

    // Get the ProjectItem associated with this sequence
    const projectItem = await sequence.getProjectItem();

    if (!projectItem) {
        throw new Error("renameSequence: Could not get ProjectItem for sequence");
    }

    // Execute rename action
    execute(() => {
        const action = projectItem.createSetNameAction(newName);
        return [action];
    }, project);

    return {
        id: sequenceId,
        newName: newName
    };
}

/**
 * Parse resolution (width x height) from project columns metadata JSON.
 * Looks for patterns like "1920 x 1080" in column values.
 */
const parseResolutionFromMetadata = (jsonMetadata) => {
    try {
        const metadata = JSON.parse(jsonMetadata);
        for (const col of metadata) {
            const value = col.ColumnValue || '';
            // Match patterns like "1920 x 1080" or "1920x1080"
            const match = value.match(/(\d+)\s*x\s*(\d+)/i);
            if (match) {
                return { width: parseInt(match[1], 10), height: parseInt(match[2], 10) };
            }
        }
    } catch (e) {
        // Parse failed, return zeros
    }
    return { width: 0, height: 0 };
};

/**
 * Parse duration from project columns metadata JSON.
 * Looks for Duration column with format like "00:00:05:15" (HH:MM:SS:FF) or seconds.
 * Returns duration in ticks.
 */
const parseDurationFromMetadata = (jsonMetadata, fps) => {
    try {
        const metadata = JSON.parse(jsonMetadata);
        for (const col of metadata) {
            const name = (col.ColumnName || '').toLowerCase();
            const value = col.ColumnValue || '';

            if (name.includes('duration') || name.includes('media duration')) {
                // Try to parse timecode format HH:MM:SS:FF or HH;MM;SS;FF
                const tcMatch = value.match(/(\d{2})[;:](\d{2})[;:](\d{2})[;:](\d{2})/);
                if (tcMatch) {
                    const hours = parseInt(tcMatch[1], 10);
                    const mins = parseInt(tcMatch[2], 10);
                    const secs = parseInt(tcMatch[3], 10);
                    const frames = parseInt(tcMatch[4], 10);
                    const totalSeconds = hours * 3600 + mins * 60 + secs + (fps > 0 ? frames / fps : 0);
                    return Math.round(totalSeconds * 254016000000);
                }
                // Try to parse as seconds (but NOT if it looks like a large number/ticks)
                const secMatch = value.match(/^(\d+(?:\.\d+)?)\s*(?:s|sec)?$/i);
                if (secMatch) {
                    const secValue = parseFloat(secMatch[1]);
                    // Sanity check: if > 86400 (1 day in seconds), it's probably ticks not seconds
                    if (secValue < 86400) {
                        return Math.round(secValue * 254016000000);
                    }
                }
            }
        }
    } catch (e) {
        // Parse failed
    }
    return 0;
};

/**
 * Parse hasVideo/hasAudio from project columns metadata JSON.
 * Looks for Video Info and Audio Info columns.
 */
const parseStreamsFromMetadata = (jsonMetadata) => {
    let hasVideo = false;
    let hasAudio = false;
    try {
        const metadata = JSON.parse(jsonMetadata);
        for (const col of metadata) {
            const name = (col.ColumnName || '').toLowerCase();
            const value = (col.ColumnValue || '');

            // Video Info column has resolution info if video exists
            if (name.includes('video') && value && value.toLowerCase() !== 'none' && value !== '-') {
                hasVideo = true;
            }
            // Audio Info column has sample rate/channels if audio exists
            if (name.includes('audio') && value && value.toLowerCase() !== 'none' && value !== '-') {
                hasAudio = true;
            }
        }
    } catch (e) {
        // Default to true if we can't parse (safer assumption)
        hasVideo = true;
        hasAudio = true;
    }
    return { hasVideo, hasAudio };
};

const getProjectItemMetadata = async (command) => {
    const options = command.options;
    const itemName = options.itemName;

    const project = await app.Project.getActiveProject();
    const item = await findProjectItem(itemName, project);

    if (!item) {
        throw new Error(`getProjectItemMetadata: Could not find project item: ${itemName}`);
    }

    // Cast to ClipProjectItem to access media-specific methods
    const clipItem = app.ClipProjectItem.cast(item);
    if (!clipItem) {
        throw new Error(`getProjectItemMetadata: Item "${itemName}" is not a media clip`);
    }

    // Get duration - try multiple approaches
    let durationTicks = 0;

    // Approach 1: Try getMedia().duration (note: duration is a Promise!)
    try {
        const media = await clipItem.getMedia();
        if (media) {
            const duration = await media.duration;  // duration is a Promise!
            if (duration) {
                // Try different properties of TickTime
                if (typeof duration.ticksNumber === 'number' && duration.ticksNumber > 0) {
                    durationTicks = duration.ticksNumber;
                } else if (duration.ticks) {
                    durationTicks = parseInt(duration.ticks, 10);
                } else if (typeof duration.seconds === 'number' && duration.seconds > 0) {
                    // Sanity check: seconds should be < 86400 (1 day)
                    if (duration.seconds < 86400) {
                        durationTicks = Math.round(duration.seconds * 254016000000);
                    } else {
                        durationTicks = Math.round(duration.seconds);
                    }
                }
            }
        }
    } catch (e) {
        // getMedia failed, will try fallback approaches
    }

    // Approach 2: Try getOutPoint (source duration = out point if in point is 0)
    if (durationTicks === 0) {
        try {
            const outPoint = await clipItem.getOutPoint();
            if (outPoint) {
                if (typeof outPoint.ticksNumber === 'number') {
                    durationTicks = outPoint.ticksNumber;
                } else if (outPoint.ticks) {
                    durationTicks = parseInt(outPoint.ticks, 10);
                }
            }
        } catch (e2) {
            // getOutPoint failed, will try MediaType.VIDEO approach
        }
    }

    // Approach 3: Try with MediaType.VIDEO
    if (durationTicks === 0) {
        try {
            const outPoint = await clipItem.getOutPoint(constants.MediaType.VIDEO);
            const inPoint = await clipItem.getInPoint(constants.MediaType.VIDEO);
            if (outPoint && inPoint) {
                const outTicks = typeof outPoint.ticksNumber === 'number' ? outPoint.ticksNumber : parseInt(outPoint.ticks, 10);
                const inTicks = typeof inPoint.ticksNumber === 'number' ? inPoint.ticksNumber : parseInt(inPoint.ticks, 10);
                durationTicks = outTicks - inTicks;
            }
        } catch (e3) {
            // getInPoint/getOutPoint(VIDEO) failed, will try metadata approach
        }
    }

    // Get fps from footage interpretation
    let fps = 0;
    try {
        const interpretation = await clipItem.getFootageInterpretation();
        if (interpretation) {
            fps = interpretation.getFrameRate();
        }
    } catch (e) {
        // Could not get fps
    }

    // Get width/height, duration, and stream info from project columns metadata
    let width = 0;
    let height = 0;
    let hasVideo = true;
    let hasAudio = true;
    let metadataJson = null;

    try {
        metadataJson = await app.Metadata.getProjectColumnsMetadata(item);

        const resolution = parseResolutionFromMetadata(metadataJson);
        width = resolution.width;
        height = resolution.height;

        const streams = parseStreamsFromMetadata(metadataJson);
        hasVideo = streams.hasVideo;
        hasAudio = streams.hasAudio;

        // Approach 4: Try to get duration from metadata if still 0
        if (durationTicks === 0 && metadataJson) {
            durationTicks = parseDurationFromMetadata(metadataJson, fps);
        }
    } catch (e) {
        // Could not get metadata
    }

    return {
        durationTicks,
        fps,
        width,
        height,
        hasVideo,
        hasAudio
    };
}

const setProjectItemInOutPoints = async (command) => {
    const options = command.options;
    const itemName = options.itemName;
    const inPointTicks = options.inPointTicks;
    const outPointTicks = options.outPointTicks;

    const project = await app.Project.getActiveProject();
    const item = await findProjectItem(itemName, project);

    if (!item) {
        throw new Error(`setProjectItemInOutPoints: Could not find project item: ${itemName}`);
    }

    // Cast to ClipProjectItem to access in/out point methods
    const clipItem = app.ClipProjectItem.cast(item);
    if (!clipItem) {
        throw new Error(`setProjectItemInOutPoints: Item "${itemName}" is not a media clip`);
    }

    const inTick = app.TickTime.createWithTicks(inPointTicks.toString());
    const outTick = app.TickTime.createWithTicks(outPointTicks.toString());

    // Use combined createSetInOutPointsAction instead of separate actions
    execute(() => {
        const action = clipItem.createSetInOutPointsAction(inTick, outTick);
        return [action];
    }, project);
}

/**
 * Get the number of video tracks in a sequence.
 *
 * @param {Object} command
 * @param {string} command.options.sequenceId - Sequence ID
 * @returns {number} Number of video tracks
 */
const getVideoTrackCount = async (command) => {
    const options = command.options;
    const sequenceId = options.sequenceId;

    const sequence = await _getSequenceFromId(sequenceId);

    if (!sequence) {
        throw new Error(`getVideoTrackCount: sequence with id [${sequenceId}] not found.`);
    }

    const count = await sequence.getVideoTrackCount();
    return { count };
};

/**
 * Check if a track is occupied in a given time range.
 * Returns true if any clip on the track overlaps the specified range.
 *
 * @param {Object} command
 * @param {string} command.options.sequenceId - Sequence ID
 * @param {number} command.options.trackIndex - Video track index to check
 * @param {number} command.options.startTicks - Start of range in ticks
 * @param {number} command.options.endTicks - End of range in ticks
 * @returns {boolean} True if any clip overlaps the range
 */
const isTrackOccupiedInRange = async (command) => {
    const options = command.options;
    const sequenceId = options.sequenceId;
    const trackIndex = options.trackIndex;
    const startTicks = BigInt(options.startTicks);
    const endTicks = BigInt(options.endTicks);

    const sequence = await _getSequenceFromId(sequenceId);

    if (!sequence) {
        throw new Error(`isTrackOccupiedInRange: sequence with id [${sequenceId}] not found.`);
    }

    // Check if track exists
    const trackCount = await sequence.getVideoTrackCount();
    if (trackIndex >= trackCount) {
        // Track doesn't exist yet - not occupied
        return { occupied: false };
    }

    const track = await sequence.getVideoTrack(trackIndex);
    if (!track) {
        return { occupied: false };
    }

    // Get all clips on track (TrackItemType 1 = CLIP)
    const clips = await track.getTrackItems(1, false);

    if (!clips || clips.length === 0) {
        return { occupied: false };
    }

    // Check each clip for overlap
    for (const clip of clips) {
        const clipStartTime = await clip.getStartTime();
        const clipEndTime = await clip.getEndTime();

        const clipStart = BigInt(clipStartTime.ticksNumber);
        const clipEnd = BigInt(clipEndTime.ticksNumber);

        // Overlap check: (clipStart < endTicks) && (clipEnd > startTicks)
        if (clipStart < endTicks && clipEnd > startTicks) {
            return { occupied: true };
        }
    }

    return { occupied: false };
};

/**
 * Find the index of a clip at a specific timeline position on a track.
 * Used to identify newly placed clips for effect copying.
 *
 * @param {Object} command
 * @param {string} command.options.sequenceId - Sequence ID
 * @param {number} command.options.trackIndex - Track index to search
 * @param {number} command.options.positionTicks - Timeline position in ticks
 * @param {string} command.options.trackType - "VIDEO" or "AUDIO"
 * @returns {number|null} Clip index if found, null otherwise
 */
const findClipIndexAtPosition = async (command) => {
    const options = command.options;
    const sequenceId = options.sequenceId;
    const trackIndex = options.trackIndex;
    const positionTicks = BigInt(options.positionTicks);
    const trackType = options.trackType || "VIDEO";

    const sequence = await _getSequenceFromId(sequenceId);

    if (!sequence) {
        throw new Error(`findClipIndexAtPosition: sequence with id [${sequenceId}] not found.`);
    }

    let track;
    if (trackType === "VIDEO") {
        track = await sequence.getVideoTrack(trackIndex);
    } else {
        track = await sequence.getAudioTrack(trackIndex);
    }

    if (!track) {
        return { clipIndex: null };
    }

    const clips = await track.getTrackItems(1, false);

    // Find clip whose start time matches the position
    for (let i = 0; i < clips.length; i++) {
        const clip = clips[i];
        const clipStartTime = await clip.getStartTime();
        const clipStart = BigInt(clipStartTime.ticksNumber);

        // Check if this clip starts at our position (allow small tolerance)
        if (clipStart === positionTicks) {
            return { clipIndex: i };
        }
    }

    return { clipIndex: null };
};

/**
 * Copy attributes (effects, motion, opacity) from one clip to another.
 * Replicates Edit > Paste Attributes functionality.
 *
 * BATCHED VERSION: Collects all actions and executes in single transaction.
 *
 * @param {Object} command
 * @param {string} command.options.sequenceId - Sequence ID
 * @param {number} command.options.sourceTrackIndex - Source clip's video track index
 * @param {number} command.options.sourceTrackItemIndex - Source clip's index within track
 * @param {number} command.options.targetTrackIndex - Target clip's video track index
 * @param {number} command.options.targetTrackItemIndex - Target clip's index within track
 * @param {boolean} command.options.copyMotion - Copy motion attributes (default: true)
 * @param {boolean} command.options.copyOpacity - Copy opacity attributes (default: true)
 * @param {boolean} command.options.copyTimeRemapping - Copy time remapping (default: false)
 * @param {boolean} command.options.copyEffects - Copy video effects (default: true)
 */
const copyClipAttributes = async (command) => {
    const options = command.options;
    const sequenceId = options.sequenceId;
    const sourceTrackIndex = options.sourceTrackIndex;
    const sourceTrackItemIndex = options.sourceTrackItemIndex;
    const targetTrackIndex = options.targetTrackIndex;
    const targetTrackItemIndex = options.targetTrackItemIndex;

    // Default options
    const copyMotion = options.copyMotion !== false;
    const copyOpacity = options.copyOpacity !== false;
    const copyTimeRemapping = options.copyTimeRemapping === true; // Off by default
    const copyEffects = options.copyEffects !== false;

    const project = await app.Project.getActiveProject();
    const sequence = await _getSequenceFromId(sequenceId);

    if (!sequence) {
        throw new Error(`copyClipAttributes: sequence with id [${sequenceId}] not found.`);
    }

    // Get source and target track items
    const sourceItem = await getTrack(sequence, sourceTrackIndex, sourceTrackItemIndex, TRACK_TYPE.VIDEO);
    const targetItem = await getTrack(sequence, targetTrackIndex, targetTrackItemIndex, TRACK_TYPE.VIDEO);

    // Get component chains
    const sourceChain = await sourceItem.getComponentChain();
    let targetChain = await targetItem.getComponentChain();

    // Intrinsic effects (always exist on clips - don't create, just copy values)
    const INTRINSIC_MOTION = 'AE.ADBE Motion';
    const INTRINSIC_OPACITY = 'AE.ADBE Opacity';
    const INTRINSIC_TIME_REMAP = 'AE.ADBE Time Remapping';

    // Track which effects to skip creating (intrinsics)
    const intrinsicEffects = new Set([INTRINSIC_MOTION, INTRINSIC_OPACITY, INTRINSIC_TIME_REMAP]);

    const copiedEffects = [];
    const copiedParams = [];
    const debugLog = [];

    // Collect all param set actions for batched execution
    const allParamActions = [];

    // Log source components for debugging
    const sourceCount = sourceChain.getComponentCount();
    debugLog.push(`Source clip has ${sourceCount} components`);

    for (let i = 0; i < sourceCount; i++) {
        const comp = sourceChain.getComponentAtIndex(i);
        const matchName = await comp.getMatchName();
        const paramCount = comp.getParamCount();
        debugLog.push(`  [${i}] ${matchName} (${paramCount} params)`);
    }

    // Helper: find component by matchName in a chain
    const findComponent = async (chain, matchName) => {
        const count = chain.getComponentCount();
        for (let i = 0; i < count; i++) {
            const comp = chain.getComponentAtIndex(i);
            const compMatchName = await comp.getMatchName();
            if (compMatchName === matchName) {
                return comp;
            }
        }
        return null;
    };

    // Helper: collect param copy actions (doesn't execute, just collects)
    // IMPORTANT: Uses INDEX-based matching, not name-based, because some components
    // have multiple params with the same displayName (e.g., Lumetri has "Temperature" twice)
    const collectParamActions = async (sourceComp, targetComp, componentName, skipDuplicates = false) => {
        const sourceParamCount = sourceComp.getParamCount();
        const targetParamCount = targetComp.getParamCount();
        debugLog.push(`Collecting params from ${componentName}: source=${sourceParamCount}, target=${targetParamCount}`);
        const zeroTime = app.TickTime.createWithTicks("0");

        // Use the minimum of source and target param counts
        const paramCount = Math.min(sourceParamCount, targetParamCount);

        // Track seen param names to skip duplicates (for intrinsic effects like Opacity)
        const seenParamNames = new Set();

        for (let p = 0; p < paramCount; p++) {
            const sourceParam = sourceComp.getParam(p);
            const targetParam = targetComp.getParam(p);
            const paramName = sourceParam.displayName;
            const targetParamName = targetParam.displayName;

            // Skip unnamed params
            if (!paramName || paramName.trim() === '') continue;

            // For intrinsic effects (Opacity, Motion), skip duplicate param names
            // This prevents the second "Blend Mode" (internal state) from overwriting the first
            if (skipDuplicates && seenParamNames.has(paramName)) {
                debugLog.push(`  [${p}] "${paramName}" - SKIP (duplicate in intrinsic)`);
                continue;
            }
            seenParamNames.add(paramName);

            // Handle uniform Scale -> Scale Width + Scale Height mapping BEFORE mismatch check
            // This is a special case where source has "Scale" but target might have separate Width/Height
            if (paramName === 'Scale' && targetParamName !== 'Scale') {
                debugLog.push(`    → uniform Scale mapping to Width+Height`);
                // Find Scale Width and Scale Height on target by name (special case)
                let scaleWidthParam = null;
                let scaleHeightParam = null;
                for (let tp = 0; tp < targetParamCount; tp++) {
                    const tp_param = targetComp.getParam(tp);
                    if (tp_param.displayName === 'Scale Width') scaleWidthParam = tp_param;
                    if (tp_param.displayName === 'Scale Height') scaleHeightParam = tp_param;
                }
                if (scaleWidthParam && scaleHeightParam) {
                    try {
                        let rawValue;
                        if (typeof sourceParam.getValueAtTime === 'function') {
                            rawValue = await sourceParam.getValueAtTime(zeroTime);
                        } else if (typeof sourceParam.getValue === 'function') {
                            rawValue = await sourceParam.getValue();
                        }
                        if (rawValue !== undefined && rawValue !== null) {
                            let value = rawValue;
                            if (rawValue && typeof rawValue === 'object' && 'value' in rawValue) {
                                value = rawValue.value;
                            }
                            debugLog.push(`    Scale value=${value}, applying to Width+Height`);
                            const widthKeyframe = await scaleWidthParam.createKeyframe(value);
                            const widthAction = scaleWidthParam.createSetValueAction(widthKeyframe);
                            allParamActions.push(widthAction);
                            copiedParams.push(`${componentName}.Scale Width`);
                            const heightKeyframe = await scaleHeightParam.createKeyframe(value);
                            const heightAction = scaleHeightParam.createSetValueAction(heightKeyframe);
                            allParamActions.push(heightAction);
                            copiedParams.push(`${componentName}.Scale Height`);
                        }
                    } catch (e) {
                        debugLog.push(`    Scale mapping ERROR: ${e.message || e}`);
                    }
                }
                continue;
            }

            // Verify names match (sanity check for index-based matching)
            if (paramName !== targetParamName) {
                debugLog.push(`  [${p}] MISMATCH: source="${paramName}" vs target="${targetParamName}" - SKIP`);
                continue;
            }

            debugLog.push(`  [${p}] "${paramName}"`);

            // Check if source param has keyframes (is time-varying)
            try {
                const isTimeVarying = sourceParam.isTimeVarying ? sourceParam.isTimeVarying() : false;

                if (isTimeVarying && typeof sourceParam.getKeyframeListAsTickTimes === 'function') {
                    // KEYFRAME MODE: Copy all keyframes
                    const keyframeTimes = await sourceParam.getKeyframeListAsTickTimes();
                    debugLog.push(`    → ${keyframeTimes.length} keyframe(s)`);

                    // CRITICAL: First enable time-varying mode on target param
                    // Without this, keyframes won't be added properly
                    if (keyframeTimes.length > 0) {
                        const setTimeVaryingAction = targetParam.createSetTimeVaryingAction(true);
                        allParamActions.push(setTimeVaryingAction);
                        debugLog.push(`      Enabled time-varying mode`);
                    }

                    // Collect interpolation actions to apply AFTER all keyframes are added
                    const interpolationActions = [];

                    // Note: Interpolation mode copying is not currently supported for intrinsic effects
                    // The UXP findPreviousKeyframe/findNextKeyframe APIs have issues with Motion/Opacity
                    // Keyframes will copy with correct values and positions but default to LINEAR interpolation

                    for (let kfIdx = 0; kfIdx < keyframeTimes.length; kfIdx++) {
                        const kfTime = keyframeTimes[kfIdx];
                        try {
                            // Interpolation mode not copied (UXP API limitation with intrinsic effects)
                            let interpMode = null;

                            // Get value at this keyframe time
                            let rawValue = await sourceParam.getValueAtTime(kfTime);
                            if (rawValue === undefined || rawValue === null) continue;

                            // Extract actual value
                            let value = rawValue;
                            if (rawValue && typeof rawValue === 'object' && 'value' in rawValue) {
                                value = rawValue.value;
                            }

                            // Convert arrays to PointF
                            if (Array.isArray(value) && value.length === 2) {
                                value = new app.PointF(value[0], value[1]);
                            }

                            // Create keyframe with value
                            const newKeyframe = await targetParam.createKeyframe(value);
                            // Set keyframe position
                            newKeyframe.position = kfTime;

                            // Add keyframe action
                            const addAction = targetParam.createAddKeyframeAction(newKeyframe);
                            allParamActions.push(addAction);

                            // Queue interpolation action to apply after keyframe is added
                            if (interpMode !== null && typeof targetParam.createSetInterpolationAtKeyframeAction === 'function') {
                                const interpAction = targetParam.createSetInterpolationAtKeyframeAction(kfTime, interpMode, false);
                                interpolationActions.push(interpAction);
                            }

                            const interpStr = interpMode !== null ? ` interp=${interpMode}` : '';
                            const timeStr = kfTime.seconds ? kfTime.seconds.toFixed(2) + 's' : (kfTime.ticks || kfTime);
                            debugLog.push(`      KF@${timeStr}: ${JSON.stringify(value).slice(0, 30)}${interpStr}`);
                        } catch (kfErr) {
                            debugLog.push(`      KF ERROR: ${kfErr.message || kfErr}`);
                        }
                    }

                    // Add interpolation actions after all keyframe adds
                    allParamActions.push(...interpolationActions);
                    copiedParams.push(`${componentName}.${paramName} (${keyframeTimes.length} KFs)`);
                } else {
                    // STATIC VALUE MODE: Copy single value at time 0
                    let rawValue;

                    // Try getValueAtTime first, then getStartValue as fallback
                    if (typeof sourceParam.getValueAtTime === 'function') {
                        try {
                            rawValue = await sourceParam.getValueAtTime(zeroTime);
                        } catch (valueErr) {
                            // getValueAtTime failed - try getStartValue as fallback
                            if (typeof sourceParam.getStartValue === 'function') {
                                try {
                                    const startKeyframe = await sourceParam.getStartValue();
                                    if (startKeyframe && startKeyframe.value !== undefined) {
                                        rawValue = startKeyframe.value;
                                    } else {
                                        debugLog.push(`    SKIP (getStartValue returned no value)`);
                                        continue;
                                    }
                                } catch (startErr) {
                                    debugLog.push(`    SKIP (complex type, no UXP copy API)`);
                                    continue;
                                }
                            } else {
                                debugLog.push(`    SKIP (complex type, no UXP copy API)`);
                                continue;
                            }
                        }
                    } else if (typeof sourceParam.getValue === 'function') {
                        rawValue = await sourceParam.getValue();
                    } else {
                        continue;
                    }

                    if (rawValue === undefined || rawValue === null) continue;

                    // Extract the actual value - getValueAtTime returns {value: X}
                    let value = rawValue;
                    if (rawValue && typeof rawValue === 'object' && 'value' in rawValue) {
                        value = rawValue.value;
                    }

                    // Handle array values (Position, Anchor Point) - convert to PointF
                    if (Array.isArray(value)) {
                        if (value.length === 2) {
                            try {
                                const pointF = new app.PointF(value[0], value[1]);
                                debugLog.push(`    [${value[0]}, ${value[1]}] → PointF`);
                                const keyframe = await targetParam.createKeyframe(pointF);
                                const action = targetParam.createSetValueAction(keyframe);
                                allParamActions.push(action);
                                copiedParams.push(`${componentName}.${paramName}`);
                            } catch (pointErr) {
                                debugLog.push(`    PointF failed - ${pointErr.message || pointErr}`);
                            }
                        } else {
                            debugLog.push(`    skipping array (length ${value.length}, not 2D)`);
                        }
                        continue;
                    }

                    debugLog.push(`    ${JSON.stringify(value)}`);

                    const keyframe = await targetParam.createKeyframe(value);
                    const action = targetParam.createSetValueAction(keyframe);
                    allParamActions.push(action);
                    copiedParams.push(`${componentName}.${paramName}`);
                }
            } catch (e) {
                debugLog.push(`    ERROR - ${e.message || e}`);
            }
        }
    };

    // PHASE 1: Create any missing non-intrinsic effects on target
    if (copyEffects) {
        const effectsToCreate = [];
        for (let i = 0; i < sourceCount; i++) {
            const sourceComp = sourceChain.getComponentAtIndex(i);
            const matchName = await sourceComp.getMatchName();

            if (intrinsicEffects.has(matchName)) continue;

            const existsOnTarget = await findComponent(targetChain, matchName);
            if (!existsOnTarget) {
                effectsToCreate.push(matchName);
            }
        }

        if (effectsToCreate.length > 0) {
            debugLog.push(`Creating ${effectsToCreate.length} effects on target...`);

            // Create all effects in one transaction
            const createActions = [];
            for (const matchName of effectsToCreate) {
                try {
                    const newEffect = await app.VideoFilterFactory.createComponent(matchName);
                    createActions.push(targetChain.createAppendComponentAction(newEffect, 0));
                    debugLog.push(`  Will create: ${matchName}`);
                } catch (e) {
                    debugLog.push(`  ERROR creating ${matchName}: ${e.message || e}`);
                }
            }

            if (createActions.length > 0) {
                execute(() => createActions, project);
                // Give Premiere time to fully initialize the new effects
                await new Promise(resolve => setTimeout(resolve, 100));
                // Re-fetch target chain after creating effects
                targetChain = await targetItem.getComponentChain();
                debugLog.push(`Effects created, refreshed target chain`);
            }
        }
    }

    // PHASE 2: Collect all param copy actions

    // Copy Motion parameters (if enabled)
    // skipDuplicates=true for intrinsic effects (prevents internal state params from overwriting)
    if (copyMotion) {
        const sourceMotion = await findComponent(sourceChain, INTRINSIC_MOTION);
        const targetMotion = await findComponent(targetChain, INTRINSIC_MOTION);
        if (sourceMotion && targetMotion) {
            await collectParamActions(sourceMotion, targetMotion, 'Motion', true);
            copiedEffects.push('Motion');
        }
    }

    // Copy Opacity parameters (if enabled)
    // skipDuplicates=true for intrinsic effects (prevents second Blend Mode from overwriting)
    if (copyOpacity) {
        const sourceOpacity = await findComponent(sourceChain, INTRINSIC_OPACITY);
        const targetOpacity = await findComponent(targetChain, INTRINSIC_OPACITY);
        if (sourceOpacity && targetOpacity) {
            await collectParamActions(sourceOpacity, targetOpacity, 'Opacity', true);
            copiedEffects.push('Opacity');
        }
    }

    // Copy Time Remapping (if enabled - off by default)
    // skipDuplicates=true for intrinsic effects
    if (copyTimeRemapping) {
        const sourceTimeRemap = await findComponent(sourceChain, INTRINSIC_TIME_REMAP);
        const targetTimeRemap = await findComponent(targetChain, INTRINSIC_TIME_REMAP);
        if (sourceTimeRemap && targetTimeRemap) {
            await collectParamActions(sourceTimeRemap, targetTimeRemap, 'Time Remapping', true);
            copiedEffects.push('Time Remapping');
        }
    }

    // Copy non-intrinsic effect parameters (if enabled)
    if (copyEffects) {
        for (let i = 0; i < sourceCount; i++) {
            const sourceComp = sourceChain.getComponentAtIndex(i);
            const matchName = await sourceComp.getMatchName();

            if (intrinsicEffects.has(matchName)) continue;

            const targetComp = await findComponent(targetChain, matchName);
            if (targetComp) {
                await collectParamActions(sourceComp, targetComp, matchName);
                copiedEffects.push(matchName);
            } else {
                debugLog.push(`WARNING: ${matchName} not found on target after creation`);
            }
        }
    }

    // PHASE 3: Execute all param actions in single batch
    debugLog.push(`Executing ${allParamActions.length} param actions in single batch...`);
    if (allParamActions.length > 0) {
        execute(() => allParamActions, project);
    }

    return {
        copiedEffects,
        copiedParams,
        debugLog,
        message: `Copied ${copiedEffects.length} effects with ${copiedParams.length} params`
    };
};

/**
 * Remove a non-intrinsic effect from a clip by matchName.
 *
 * @param {Object} command
 * @param {string} command.options.sequenceId - Sequence ID
 * @param {number} command.options.trackIndex - Video track index
 * @param {number} command.options.trackItemIndex - Clip index within track
 * @param {string} command.options.effectMatchName - The matchName of the effect to remove (e.g., "AE.ADBE Lumetri")
 * @param {boolean} command.options.removeAll - If true, remove all instances of this effect (default: false, removes first found)
 */
const removeEffectFromClip = async (command) => {
    const options = command.options;
    const sequenceId = options.sequenceId;
    const trackIndex = options.trackIndex;
    const trackItemIndex = options.trackItemIndex;
    const effectMatchName = options.effectMatchName;
    const removeAll = options.removeAll === true;

    const project = await app.Project.getActiveProject();
    const sequence = await _getSequenceFromId(sequenceId);

    if (!sequence) {
        throw new Error(`removeEffectFromClip: sequence with id [${sequenceId}] not found.`);
    }

    const trackItem = await getTrack(sequence, trackIndex, trackItemIndex, TRACK_TYPE.VIDEO);
    const componentChain = await trackItem.getComponentChain();
    const count = componentChain.getComponentCount();

    const componentsToRemove = [];

    for (let i = 0; i < count; i++) {
        const comp = componentChain.getComponentAtIndex(i);
        const matchName = await comp.getMatchName();
        if (matchName === effectMatchName) {
            componentsToRemove.push(comp);
            if (!removeAll) break;
        }
    }

    if (componentsToRemove.length === 0) {
        return {
            removed: 0,
            message: `No effect with matchName "${effectMatchName}" found on clip`
        };
    }

    execute(() => {
        const actions = [];
        for (const comp of componentsToRemove) {
            actions.push(componentChain.createRemoveComponentAction(comp));
        }
        return actions;
    }, project);

    return {
        removed: componentsToRemove.length,
        message: `Removed ${componentsToRemove.length} instance(s) of "${effectMatchName}"`
    };
};

/**
 * Get the effects and their parameter values from a clip.
 * Used to verify if effects/params were applied correctly.
 *
 * @param {Object} command
 * @param {string} command.options.sequenceId - Sequence ID
 * @param {number} command.options.trackIndex - Video track index
 * @param {number} command.options.trackItemIndex - Clip index within track
 */
const getClipEffects = async (command) => {
    const options = command.options;
    const sequenceId = options.sequenceId;
    const trackIndex = options.trackIndex;
    const trackItemIndex = options.trackItemIndex;

    const sequence = await _getSequenceFromId(sequenceId);

    if (!sequence) {
        throw new Error(`getClipEffects: sequence with id [${sequenceId}] not found.`);
    }

    const trackItem = await getTrack(sequence, trackIndex, trackItemIndex, TRACK_TYPE.VIDEO);
    const componentChain = await trackItem.getComponentChain();
    const count = componentChain.getComponentCount();

    const effects = [];
    const zeroTime = app.TickTime.createWithTicks("0");

    for (let i = 0; i < count; i++) {
        const comp = componentChain.getComponentAtIndex(i);
        const matchName = await comp.getMatchName();
        const paramCount = comp.getParamCount();

        const params = [];
        for (let p = 0; p < paramCount; p++) {
            const param = comp.getParam(p);
            const paramName = param.displayName;

            let value = null;
            try {
                if (typeof param.getValueAtTime === 'function') {
                    const rawValue = await param.getValueAtTime(zeroTime);
                    if (rawValue && typeof rawValue === 'object' && 'value' in rawValue) {
                        value = rawValue.value;
                    } else {
                        value = rawValue;
                    }
                }
            } catch (e) {
                // Some params don't support getValueAtTime
                value = `<error: ${e.message || e}>`;
            }

            // Only include params with names (skip unnamed ones)
            if (paramName && paramName.trim() !== '') {
                params.push({
                    name: paramName,
                    value: value
                });
            }
        }

        effects.push({
            index: i,
            matchName: matchName,
            paramCount: paramCount,
            params: params // Return all params (Lumetri has 130)
        });
    }

    return { effects };
};

const commandHandlers = {
    removeEffectFromClip,
    getClipEffects,
    copyClipAttributes,
    getVideoTrackCount,
    isTrackOccupiedInRange,
    findClipIndexAtPosition,
    exportSequence,
    exportSequenceWithProgress,
    moveProjectItemsToBin,
    batchMoveItemsToBins,
    createBinInActiveProject,
    createBinInBin,
    addMarkerToSequence,
    closeGapsOnSequence,
    removeItemFromSequence,
    setClipStartEndTimes,
    setClipInOutPoints,
    moveClip,
    openProject,
    saveProjectAs,
    saveProject,
    getProjectInfo,
    setActiveSequence,
    openSequence,
    exportFrame,
    setVideoClipProperties,
    createSequenceFromMedia,
    setAudioTrackMute,
    setClipDisabled,
    appendVideoTransition,
    appendVideoFilter,
    addMediaToSequence,
    importMedia,
    createProject,
    cloneSequence,
    renameSequence,
    getProjectItemMetadata,
    setProjectItemInOutPoints,
};

module.exports = {
    commandHandlers
}