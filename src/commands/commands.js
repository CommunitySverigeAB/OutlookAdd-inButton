/* global Office */

import { createNestablePublicClientApplication } from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: "8d7be25b-0788-44d7-8575-1fcdd6ee79c2",
    authority: "https://login.microsoftonline.com/e38b60ae-3085-4e21-9ac2-f5f62fe15c0d"
  },
};

const graphScopes = ["User.Read", "Mail.ReadWrite"];

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
      persistent,
    },
    () => {}
  );
}

async function getGraphToken() {
  const pca = await createNestablePublicClientApplication(msalConfig);

  const accounts = pca.getAllAccounts();
  if (accounts.length > 0) {
    try {
      const silent = await pca.acquireTokenSilent({
        account: accounts[0],
        scopes: graphScopes,
      });
      console.log("Token acquired silently");
      return silent.accessToken;
    } catch (err) {
      console.warn("Silent token failed, trying popup:", err);
    }
  }

  const interactive = await pca.acquireTokenPopup({
    scopes: graphScopes,
  });

  console.log("Token acquired interactively");
  return interactive.accessToken;
}

async function flagMessageWithGraph(item, accessToken) {
  const restId = Office.context.mailbox.convertToRestId(
    item.itemId,
    Office.MailboxEnums.RestVersion.v2_0
  );

  console.log("Original itemId:", item.itemId);
  console.log("Converted REST id:", restId);

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(restId)}`,
    {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        flag: {
          flagStatus: "flagged",
        },
      }),
    }
  );

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Graph PATCH failed: ${response.status} ${text}`);
  }

  console.log("Message flagged successfully in Graph");
}

function setOrderCategory(event) {
  console.log("Button clicked");

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
      "Den här Outlook-klienten stödjer inte kategorier.",
      true
    );
    event.completed();
    return;
  }

  item.categories.addAsync([categoryName], async (result) => {
    console.log("Category add result:", result);

    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      notify(
        item,
        Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        `Kunde inte sätta kategori: ${result.error.code} - ${result.error.message}`,
        true
      );
      event.completed();
      return;
    }

    try {
      const accessToken = await getGraphToken();
      await flagMessageWithGraph(item, accessToken);

      notify(
        item,
        Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        "Kategorin sattes och mailet flaggades."
      );
    } catch (err) {
      console.error("Flagging failed:", err);

      notify(
        item,
        Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        `Kategorin sattes, men flaggning misslyckades: ${err.message}`,
        true
      );
    }

    event.completed();
  });
}

Office.actions.associate("setOrderCategory", setOrderCategory);