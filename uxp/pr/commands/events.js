/**
 * Event emitter for streaming events back to clients.
 *
 * Used by long-running operations (like export) to send progress updates.
 * main.js sets the socket, handlers call sendEvent().
 */

let socket = null;

const setSocket = (s) => {
    socket = s;
};

/**
 * Send a progress/status event to the client.
 *
 * @param {string} eventType - Type of event (e.g., 'export_progress', 'export_complete')
 * @param {object} data - Event data to send
 */
const sendEvent = (eventType, data) => {
    if (socket && socket.connected) {
        socket.emit('event', {
            type: eventType,
            ...data
        });
        return true;
    }
    console.warn(`[events] Cannot send ${eventType}: socket not connected`);
    return false;
};

module.exports = {
    setSocket,
    sendEvent
};
