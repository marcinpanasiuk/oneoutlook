Office.initialize = function (reason) { };

/**
 * Handles the OnNewMessageCompose event
 */
function onNewMessageComposeHandler(event) {

    event.completed();
}

if (Office.actions) {
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
else {
    window["onNewMessageComposeHandler"] = onNewMessageComposeHandler;
}
