
function onMessageSendHandler(event) {
  event.completed({
    allowEvent: false,
    errorMessage: 'hello from harmon.ie',
    sendModeOverride: 'PromptUser',
  });
}

window.onMessageSendHandler = onMessageSendHandler;
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("OnMessageComposeHandler", onMessageSendHandler);

