// src/commands.ts

/* global Office, OfficeRuntime */

interface DialogMessage {
  action: "save" | "cancel";
  subject?: string;
}

let dialog: Office.Dialog | null = null;

/**
 * Command-Funktion, die über das Manifest aufgerufen wird, wenn
 * der Benutzer in einer bestehenden Mail auf den Add-in-Button klickt.
 */
export function showSubjectDialog(event: Office.AddinCommands.Event) {
  const item = Office.context.mailbox.item as Office.MessageRead;
  const subject = item.subject || "";

  const dialogUrl = buildDialogUrl(subject);

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    {
      height: 40,
      width: 40,
      displayInIframe: true
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        dialog = asyncResult.value;

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          onDialogMessageReceived
        );
      } else {
        console.error("Dialog konnte nicht geöffnet werden:", asyncResult.error);
      }

      // Wichtig: Dem Office-Host mitteilen, dass der Command abgeschlossen ist.
      event.completed();
    }
  );
}

/**
 * Erzeugt die URL zum Dialog und übergibt den Betreff als Query-Parameter.
 */
function buildDialogUrl(subject: string): string {
  // Hier wird davon ausgegangen, dass der Dialog unter:
  // https://localhost:3000/dialog.html
  // ausgeliefert wird. Bei anderem Hosting die URL entsprechend anpassen.
  const baseUrl = "https://localhost:3000/dialog.html";

  const url = new URL(baseUrl);
  url.searchParams.set("subject", subject);

  return url.toString();
}

/**
 * Wird aufgerufen, wenn der Dialog eine Nachricht an den Host sendet.
 */
async function onDialogMessageReceived(arg: Office.DialogMessageReceivedEventArgs) {
  try {
    const message = JSON.parse(arg.message) as DialogMessage;

    if (message.action === "save" && typeof message.subject === "string") {
      await updateSubject(message.subject);
    }

    if (dialog) {
      dialog.close();
      dialog = null;
    }
  } catch (err) {
    console.error("Fehler beim Verarbeiten der Dialognachricht:", err);
    if (dialog) {
      dialog.close();
      dialog = null;
    }
  }
}

/**
 * Aktualisiert den Betreff der aktuellen E-Mail über Microsoft Graph.
 * Dafür wird SSO über OfficeRuntime.auth.getAccessToken verwendet.
 */
async function updateSubject(newSubject: string): Promise<void> {
  const item = Office.context.mailbox.item as Office.MessageRead;
  const itemId = item.itemId;

  if (!itemId) {
    throw new Error("Kein itemId verfügbar.");
  }

  // REST-kompatible ID ermitteln (für den Graph-Aufruf).
  let restId = itemId;

  const mailbox = Office.context.mailbox as any;
  if (mailbox.convertToRestId) {
    restId = mailbox.convertToRestId(
      itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }

  // Access Token über SSO holen
  const accessToken = await OfficeRuntime.auth.getAccessToken({
    allowSignInPrompt: true
  });

  const graphUrl = `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(
    restId
  )}`;

  const response = await fetch(graphUrl, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      subject: newSubject
    })
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(
      `Fehler beim Aktualisieren des Betreffs über Graph: ${response.status} - ${text}`
    );
  }

  console.log("Betreff erfolgreich aktualisiert.");
}

// Export-Objekt für Office-kompatible Commands
// (wird im Manifest referenziert)
Office.actions.associate("showSubjectDialog", showSubjectDialog);
