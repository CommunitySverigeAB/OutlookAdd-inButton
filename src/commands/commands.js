/* global Office */

Office.onReady(() => {
  // ready
});

async function setOrderCategoryAndFlag(event) {
  try {
    const item = Office.context.mailbox.item;

    item.categories.addAsync(["Beställning"], (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        item.notificationMessages.replaceAsync(
          "orderDone",
          {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Kategorin Beställning lades till.",
            icon: "Icon.16x16",
            persistent: false
          },
          () => {
            event.completed();
          }
        );
      } else {
        item.notificationMessages.replaceAsync(
          "orderErr",
          {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: `Fel: ${res.error.message}`,
            icon: "Icon.16x16",
            persistent: false
          },
          () => {
            event.completed();
          }
        );
      }
    });
  } catch (e) {
    try {
      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "orderErr",
        {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: `Fel: ${e.message || e}`,
          icon: "Icon.16x16",
          persistent: false
        },
        () => {
          event.completed();
        }
      );
    } catch {
      event.completed();
    }
  }
}

Office.actions.associate("setOrderCategoryAndFlag", setOrderCategoryAndFlag);