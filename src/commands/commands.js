/* global Office */

Office.onReady(() => {
  // Ready
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
  const item = Office.context.mailbox.item;
  const categoryName = "Beställning";

  if (!item) {
    event.completed();
    return;
  }

  if (!item.categories) {
    notify(
      item,
      Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      "Den här Outlook-klienten stödjer inte kategorier för det här scenariot.",
      true
    );
    event.completed();
    return;
  }

  // Frivillig felsökning: visa om det faktiskt är shared mailbox/delegate-scenario.
  if (typeof item.getSharedPropertiesAsync === "function") {
    item.getSharedPropertiesAsync((sharedResult) => {
      if (sharedResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Shared mailbox detected:", sharedResult.value);
      } else {
        console.log("Not a shared mailbox/delegate item or shared properties unavailable.");
      }
    });
  }

  item.categories.addAsync([categoryName], (result) => {
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
        `Kunde inte sätta kategorin '${categoryName}'. Säkerställ att kategorin redan finns i den delade brevlådan. Tekniskt fel: ${result.error.message}`,
        true
      );
    }

    event.completed();
  });
}

Office.actions.associate("setOrderCategory", setOrderCategory);