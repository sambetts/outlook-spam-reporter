/// <reference path="../../node_modules/easyews/easyews.js" />

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {

  var request = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    '  <soap:Body>' +
    '    <m:CreateItem MessageDisposition="SendAndSaveCopy">' +
    '      <m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems" /></m:SavedItemFolderId>' +
    '      <m:Items>' +
    '        <t:Message>' +
    '          <t:Subject>Hello, Outlook!</t:Subject>' +
    '          <t:Body BodyType="HTML">Hello World!</t:Body>' +
    '          <t:ToRecipients>' +
    '            <t:Mailbox><t:EmailAddress>sambetts@outlook.com</t:EmailAddress></t:Mailbox>' +
    '          </t:ToRecipients>' +
    '        </t:Message>' +
    '      </m:Items>' +
    '    </m:CreateItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  sendSuspiciousMessage();
  Office.context.mailbox.makeEwsRequestAsync(request, (asyncResult: Office.AsyncResult<string>) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {

      const message: Office.NotificationMessageDetails = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Action failed with error: " + asyncResult.error.message,
        icon: "Icon.80x80",
        persistent: true,
      };

      // Show a notification message
      Office.context.mailbox.item.notificationMessages.replaceAsync("ActionPerformanceNotification", message);
    } else {
      const message: Office.NotificationMessageDetails = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Message sent!",
        icon: "Icon.80x80",
        persistent: true,
      };
      // Show a notification message
      Office.context.mailbox.item.notificationMessages.replaceAsync("ActionPerformanceNotification", message);
    }

    // Be sure to indicate when the add-in command function is complete
    event.completed();
  });
}

function sendSuspiciousMessage() {
  var item = Office.context.mailbox.item;
  var itemId = item.itemId;
  easyEws.getMailItemMimeContent(itemId, function (mimeContent) {
    var toAddress = "sambetts@microsoft.com";
    easyEws.sendPlainTextEmailWithAttachment("Suspicious Email Alert",
      "A user has forwarded a suspicious email",
      toAddress,
      "Suspicious_Email.eml",
      mimeContent,
      function (result) { console.log(result); },
      function (error) { console.log(error); },
      function (debug) { console.log(debug); });
  }, function (error) { console.log(error); }, function (debug) { console.log(debug); });
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
        ? global
        : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;
