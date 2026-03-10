/* global Office */

Office.onReady(() => {
  // ready
});

function setOrderCategory(event) {
  const item = Office.context.mailbox.item;

  if (!item) {
    event.completed();
    return;
  }

  item.categories.addAsync(["Beställning"], (res) => {
    const message =
      res.status === Office.AsyncResultStatus.Succeeded
        ? {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Kategorin Beställning lades till.",
            icon: "Icon.16x16",
            persistent: false
          }
        : {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: `Fel: ${res.error.message}`,
            icon: "Icon.16x16",
            persistent: false
          };

    item.notificationMessages.replaceAsync("orderStatus", message, () => {
      event.completed();
    });
  });
}

Office.actions.associate("setOrderCategory", setOrderCategory);