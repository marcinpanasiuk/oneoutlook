Office.initialize = function (reason) { };

/**
 * Handles the OnNewMessageCompose event
 */
function onNewMessageComposeHandler(event) {
    Office.context.mailbox.item.internetHeaders.setAsync({"x-custom-header": "test"}, function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Successfully added custom header");
            Office.context.mailbox.item.notificationMessages.replaceAsync("message", {
                type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
                actions: [{ actionText: 'Open add-in pane', actionType: 'showTaskPane', commandId: 'paneButton', contextData: '' }],
                message: 'Successfully added x-custom-header from autorun ',
                icon: 'icon16'
            });
        } else {
            console.log("Error adding custom header: " + JSON.stringify(asyncResult.error));
        }
        event.completed();
      });
}

if (Office.actions) {
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}
else {
    window["onNewMessageComposeHandler"] = onNewMessageComposeHandler;
}
