/* global Office, OfficeRuntime */

Office.onReady(() => {
  // noop
});

// ---- Hjälpare: lägg till kategori "Beställning" på item ----
function addOrderCategoryAsync(item) {
  return new Promise((resolve, reject) => {
    item.categories.addAsync(["Beställning"], (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(res.error);
      }
    });
  });
}

// ---- Hjälpare: säkerställ att "Beställning" finns i masterlistan (kräver ReadWrteMailbox) ----
function ensureOrderCategoryInMasterAsync(mailbox) {
  return new Promise((resolve) => {
    mailbox.masterCategories.getAsync((getRes) => {
      if (getRes.status !== Office.AsyncResultStatus.Succeeded) {
        // Om vi inte kan läsa – försök ändå lägga till på item
        return resolve();
      }
      const existing = (getRes.value || []).some(c => (c.displayName || c) === "Beställning");
      if (existing) return resolve();

      const newCat = [{ displayName: "Beställning", color: Office.MailboxEnums.CategoryColor.Preset0 }];
      mailbox.masterCategories.addAsync(newCat, (addRes) => {
        // Oavsett utfall – fortsätt
        resolve();
      });
    });
  });
}

// ---- Hjälpare: hämta Graph token via SSO ----
async function getGraphTokenAsync() {
  // Kräver WebApplicationInfo i manifestet
  return Office.auth.getAccessToken({ allowSignInPrompt: true });
}

// ---- Hjälpare: PATCH flag -> flagged ----
async function flagMessageAsync(item) {
  const itemId = item.itemId; // EWS/REST id, funkar i nya Outlook/OWA
  if (!itemId) throw new Error("Saknar itemId i det här läget.");

  // Identifiera om det är delad brevlåda (kräver Mailbox 1.11+ i klienten).
  let usersPath = "me";
  if (item.getSharedPropertiesAsync) {
    const sp = await new Promise((resolve, reject) => {
      item.getSharedPropertiesAsync((r) => r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error));
    });
    if (sp && sp.owner) {
      // Ägaren (SMTP) av det item du läser – använd /users/{owner}
      usersPath = `users/${encodeURIComponent(sp.owner)}`;
    }
  }

  const token = await getGraphTokenAsync();
  const resp = await fetch(`https://graph.microsoft.com/v1.0/${usersPath}/messages/${encodeURIComponent(itemId)}`, {
    method: "PATCH",
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ flag: { flagStatus: "flagged" } })
  });

  if (!resp.ok) {
    const t = await resp.text();
    throw new Error(`Graph flag misslyckades: ${resp.status} ${t}`);
  }
}

// ---- UI-less command: körs när du klickar på knappen ----
export async function setOrderCategoryAndFlag(event) {
  try {
    const { mailbox } = Office.context;
    const { item } = mailbox;

    // 1) Säkerställ kategori i masterlistan (no-op om den redan finns)
    await ensureOrderCategoryInMasterAsync(mailbox); // kräver ReadWriteMailbox enl. docs. [2](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/categories)

    // 2) Lägg till kategori på item
    await addOrderCategoryAsync(item); // [1](https://learn.microsoft.com/en-us/javascript/api/outlook/office.categories?view=outlook-js-preview)

    // 3) Flagga via Graph
    await flagMessageAsync(item); // [3](https://learn.microsoft.com/en-us/graph/api/message-update?view=graph-rest-1.0)[4](https://learn.microsoft.com/en-us/graph/api/resources/followupflag?view=graph-rest-1.0)

    // 4) Visa en liten notis i läsfönstret (valfritt)
    if (item.notificationMessages) {
      item.notificationMessages.replaceAsync("orderDone", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Kategoriserad som 'Beställning' och flaggad.",
        icon: "Icon16",
        persistent: false
      }, () => {});
    }
  } catch (e) {
    // Visa fel i en notis (valfritt)
    try {
      Office.context.mailbox.item.notificationMessages.replaceAsync("orderErr", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `Fel: ${e.message || e}`,
        icon: "Icon16",
        persistent: false
      }, () => {});
    } catch {}
  } finally {
    // Signalera att kommandot är klart
    event.completed();
  }
}