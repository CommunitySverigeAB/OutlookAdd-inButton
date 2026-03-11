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

  const mailbox = Office.context.mailbox;
  const item = mailbox.item;
  const categoryName = "Beställning";

  console.log("Mailbox:", mailbox);
  console.log("Item:", item);

  if (!item) {
    console.log("No item found");
    event.completed();
    return;
  }

  console.log("Existing categories on item:", item.categories);

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

  // 🔎 DEBUG: visa mailboxens master categories
  mailbox.masterCategories.getAsync((res) => {
    console.log("Master categories response:", res);

    if (res.status === Office.AsyncResultStatus.Succeeded) {
      const names = res.value.map(c => c.displayName);
      console.log("Master category names:", names);
    } else {
      console.error("Failed to read master categories:", res.error);
    }
  });

  // 🔎 DEBUG: visa om detta är shared mailbox
  if (typeof item.getSharedPropertiesAsync === "function") {
    item.getSharedPropertiesAsync((sharedResult) => {
      if (sharedResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Shared mailbox detected:", sharedResult.value);
      } else {
        console.log("Not a shared mailbox or shared properties unavailable");
      }
    });
  }

  console.log("Attempting to add category:", categoryName);

  item.categories.addAsync([categoryName], (result) => {
    console.log("addAsync result:", result);

    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Category added successfully");

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
}

Office.actions.associate("setOrderCategory", setOrderCategory);