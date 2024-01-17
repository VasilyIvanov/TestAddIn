/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = runMulti;

    for (let i = 1; i <= 15; i++) {
      console.log(`Req set 1.${i}`, Office.context.requirements.isSetSupported("Mailbox", `1.${i}`));
    }

    // Register an event handler to identify when messages are selected.
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, runSingle, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
      }

      console.log("ItemChanged Event handler added.");
    });

    Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, runMulti, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
      }

      console.log("SelectedItemsChanged Event handler added.");
    });
  }

  console.log("Initial item", Office.context.mailbox?.item);
});

export function runSingle() {
  console.log("Single event");

  // Clear list of previously selected messages, if any.
  const list = document.getElementById("selected-items");
  while (list.firstChild) {
    list.removeChild(list.firstChild);
  }

  console.log("Selected item", Office.context.mailbox?.item);

  const listItem = document.createElement("li");
  listItem.textContent = Office.context.mailbox?.item?.subject ?? "[NO SELECTION]";
  list.appendChild(listItem);

  console.log("Closing...");
  //Office.context.ui.closeContainer();
}

export async function runMulti() {
  // Clear list of previously selected messages, if any.
  const list = document.getElementById("selected-items");
  while (list.firstChild) {
    list.removeChild(list.firstChild);
  }

  // Retrieve the subject line of the selected messages and log it to a list in the task pane.
  //(Office.context.mailbox as any).initialData = { permissionLevel: 3 };

  Office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
    console.log("Multi event");

    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(asyncResult.error.message);
      console.log(asyncResult.error);
      return;
    }

    console.log("Selected items", asyncResult.value);

    asyncResult.value.forEach((item) => {
      const listItem = document.createElement("li");
      listItem.textContent = item.subject;
      list.appendChild(listItem);
    });
  });
}
