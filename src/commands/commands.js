/* global Office */

Office.onReady(() => {
  // Outlook add-in ready
});

function notify(item, type, message, persistent = false) {
  item.notificationMessages.replaceAsync("orderStatus", {
    type,
    message,
    icon: "Icon.16x16",
    persistent
  });
}

function setOrderCategory(event) {
  const item = Office.context.mailbox.item;
  const categoryName = "Beställning";

  if (!item) {
    event.completed();
    return;
  }

  if (!Office.context.mailbox.masterCategories) {
    notify(
      item,
      Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      "Klienten stödjer inte master categories för detta add-in.",
      true
    );
    event.completed();
    return;
  }

  const applyCategoryToItem = () => {
    item.categories.addAsync([categoryName], (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        notify(
          item,
          Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          "Kategorin Beställning lades till."
        );
      } else {
        console.error("item.categories.addAsync failed", result.error);
        notify(
          item,
          Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          `Kunde inte sätta kategori: ${result.error.message}`,
          true
        );
      }

      event.completed();
    });
  };

  Office.context.mailbox.masterCategories.getAsync((masterResult) => {
    if (masterResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("masterCategories.getAsync failed", masterResult.error);
      notify(
        item,
        Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        `Kunde inte läsa kategorier: ${masterResult.error.message}`,
        true
      );
      event.completed();
      return;
    }

    const existingCategories = masterResult.value || [];
    const categoryExists = existingCategories.some(
      (c) => c.displayName === categoryName
    );

    if (categoryExists) {
      applyCategoryToItem();
      return;
    }

    Office.context.mailbox.masterCategories.addAsync(
      [
        {
          displayName: categoryName,
          color: Office.MailboxEnums.CategoryColor.Preset0
        }
      ],
      (addResult) => {
        if (addResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("masterCategories.addAsync failed", addResult.error);
          notify(
            item,
            Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
            `Kunde inte skapa kategorin: ${addResult.error.message}`,
            true
          );
          event.completed();
          return;
        }

        applyCategoryToItem();
      }
    );
  });
}

Office.actions.associate("setOrderCategory", setOrderCategory);