/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    // Register an event handler to identify when messages are selected.
    Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, run, asyncResult => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
      }

      console.log("Event handler added.");
    });
  }
});

export async function run() {
  // Clear list of previously selected messages, if any.
  const list = document.getElementById("selected-items");
  while (list.firstChild) {
    list.removeChild(list.firstChild);
  }

  // Retrieve the subject line of the selected messages and log it to a list in the task pane.
  //(Office.context.mailbox as any).initialData = { permissionLevel: 3 };

  Office.context.mailbox.getSelectedItemsAsync(asyncResult => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(asyncResult.error.message);
      console.log(asyncResult.error);
      return;
    }

    asyncResult.value.forEach(item => {
      const listItem = document.createElement("li");
      listItem.textContent = item.subject;
      list.appendChild(listItem);
    });
  });
}
