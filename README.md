# Outlook-Web-Add-In: Betreff-Editor

Dieses Projekt enthält ein Outlook-Web-Add-In, mit dem du den Betreff einer **bestehenden E-Mail** in Outlook über ein Dialogfenster bearbeiten kannst.

## Funktionsumfang

- Das Add-In erscheint bei bestehenden E-Mails (Lesemodus) als Button im Ribbon.
- Beim Klick auf den Button:
  - Öffnet sich ein Dialogfenster.
  - Im Dialog wird der aktuelle Betreff der ausgewählten E-Mail angezeigt.
  - Der Betreff ist in einem editierbaren Textfeld änderbar.
  - Es gibt zwei Buttons:
    - **Speichern**
    - **Abbrechen**
- **Speichern**:
  - Der angepasste Betreff wird in der E-Mail über Microsoft Graph aktualisiert.
  - Das Dialogfenster wird geschlossen.
- **Abbrechen**:
  - Der Betreff wird **nicht** geändert.
  - Das Dialogfenster wird geschlossen.

## Technischer Überblick

- Sprache: **TypeScript**
- Host: **Outlook im Web** (und andere Outlook-Clients mit Add-In-Unterstützung)
- API: **Office.js** (Mailbox Requirement Set 1.8+)
- Betreff-Update:
  - Über **Microsoft Graph** (`PATCH /me/messages/{id}`).
  - Authentifizierung via **OfficeRuntime.auth.getAccessToken** (SSO).
- Projektstruktur:
  - `src/commands.ts`  
    Öffnet den Dialog und verarbeitet die Rückmeldung (Speichern/Abbrechen).
  - `src/dialog.ts`  
    Logik im Dialogfenster (UI-Ereignisse, Nachricht an den Host).
  - `public/dialog.html`  
    HTML-Dialogseite mit Eingabefeld und Buttons.
  - `manifest.xml`  
    Outlook Add-In Manifest (ExtensionPoint, Ribbon-Button, Ressourcen).

## Voraussetzungen

- Microsoft 365 Tenant mit Outlook (Web oder Desktop).
- Ein Benutzerkonto mit ausreichenden Rechten.
- Node.js (empfohlen >= 18) und npm.
- Zugriff auf eine Umgebung, in der du:
  - Das Manifest sideloaden kannst.
  - Die Webdateien über **HTTPS** hosten kannst (z.B. `https://localhost:3000`).

### AAD-App-Registrierung für Graph / SSO

Um den Betreff über Microsoft Graph ändern zu können, benötigst du eine App-Registrierung in Azure AD mit u.a.:

- API-Berechtigungen:
  - `Mail.ReadWrite` (delegiert)
- Konfiguration für SSO mit Office Add-ins:
  - Client-ID in der Add-In-Konfiguration hinterlegen (siehe Microsoft-Dokumentation zu SSO mit Office-Add-ins).

> Hinweis: In diesem Beispielprojekt ist die SSO/App-Registrierung nicht vollständig im Manifest verdrahtet.  
> Du musst die Client-ID und ggf. zusätzliche SSO-Einträge entsprechend deiner Umgebung ergänzen.

## Installation & Build

1. **Repository klonen**

   ```bash
   git clone <dein-repo-url> outlook-subject-edit-addin
   cd outlook-subject-edit-addin
