
function onMessageSendHandler(event) {
  event.completed({
    allowEvent: false,
    errorMessage: 'hello from harmon.ie',
    sendModeOverride: 'PromptUser',
  });
}

function OnMessageComposeHandler(event) {
  Office.context.mailbox.item.notificationMessages.addAsync('fd90eb33431b46f58a68720c36154b4a', {
    type: 'insightMessage',
    message: 'You can upload this email to harmon.ie',
    icon: 'Icon.16x16',
    actions: [
      {
        actionType: 'showTaskPane',
        actionText: 'Upload email to harmon.ie',
      },
    ],
  });
}

window.onMessageSendHandler = onMessageSendHandler;
window.OnMessageComposeHandler = OnMessageComposeHandler;
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("OnMessageComposeHandler", OnMessageComposeHandler);

