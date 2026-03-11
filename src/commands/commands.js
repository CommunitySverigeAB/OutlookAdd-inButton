/* global Office */

Office.onReady(() => {
  console.log("Office add-in ready");
});

function notify(item, type, message, persistent = false) {
  item.notificationMessages.replaceAsync(
    "orderStatus",
    {
      type,
      message,
      icon: "Icon.16x16",
      persistent
    },
    () => {}
  );
}

function setOrderCategory(event) {
  console.log("Button clicked");

  const item = Office.context.mailbox.item;
  const categoryName = "Beställning";

  console.log("Item:", item);

  if (!item) {
    console.log("No item found");
    event.completed();
    return;
  }

  console.log("Existing categories object:", item.categories);

  if (!item.categories) {
    console.log("item.categories missing");
    notify(
      item,
      Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      "Den här Outlook-klienten stödjer inte kategorier för det här scenariot.",
      true
    );
    event.completed();
    return;
  }

  console.log("Before addAsync");

  try {
    item.categories.addAsync([categoryName], (result) => {
      console.log("Inside addAsync callback");
      console.log("addAsync result:", result);

      if (result.status === Office.AsyncResultStatus.Succeeded) {
        notify(
          item,
          Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          "Kategorin Beställning lades till."
        );
      } else {
        console.error("item.categories.addAsync failed:", result.error);
        notify(
          item,
          Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          `Kunde inte sätta kategorin '${categoryName}'. Fel: ${result.error.code} - ${result.error.message}`,
          true
        );
      }

      event.completed();
    });
  } catch (err) {
    console.error("Synchronous error before callback:", err);
    notify(
      item,
      Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      `Ov\u00e4ntat fel: ${err.message || err}`,
      true
    );
    event.completed();
  }

  console.log("After addAsync call");
}

Office.actions.associate("setOrderCategory", setOrderCategory);