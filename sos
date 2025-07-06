(() => {
  'use strict';
  Office.actions.associate('OnMessageComposeHandler', (async function(e) {
    var t, s;
    null === (s = null === (t = Office.context.mailbox) || void 0 === t ? void 0 : t.item) || void 0 === s || s.subject.setAsync('Set by an event-based add-in!', { asyncContext: e }, (function(e) {
      e.status !== Office.AsyncResultStatus.Succeeded && console.error(`Failed to set subject: ${JSON.stringify(e.error)}`), e.asyncContext.completed();
    }));
  }));
})();
