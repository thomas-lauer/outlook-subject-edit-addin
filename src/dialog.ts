// src/dialog.ts

/* global Office */

function getSubjectFromQuery(): string {
  try {
    const params = new URLSearchParams(window.location.search);
    return params.get("subject") || "";
  } catch {
    return "";
  }
}

function init() {
  const subjectInput = document.getElementById("subject-input") as HTMLInputElement | null;
  const saveButton = document.getElementById("btn-save") as HTMLButtonElement | null;
  const cancelButton = document.getElementById("btn-cancel") as HTMLButtonElement | null;

  if (!subjectInput || !saveButton || !cancelButton) {
    console.error("Dialog-Elemente wurden nicht gefunden.");
    return;
  }

  subjectInput.value = getSubjectFromQuery();

  saveButton.addEventListener("click", () => {
    const newSubject = subjectInput.value;

    const message = {
      action: "save",
      subject: newSubject
    };

    Office.context.ui.messageParent(JSON.stringify(message));
  });

  cancelButton.addEventListener("click", () => {
    const message = {
      action: "cancel"
    };

    Office.context.ui.messageParent(JSON.stringify(message));
  });
}

Office.onReady(() => {
  init();
});
