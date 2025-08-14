/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

//import * as child_process from "child_process";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    // document.getElementById("file").onclick = getBase64;
    document.getElementById("file").onchange = getBase64;
    document.getElementById("insert-doc").onclick = insertDocument;

    document.getElementById("insertContentControl").onclick = insertContentControl;
    document.getElementById("save").onclick = save;
  }

  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("protectSheet").onclick = protectSheet;
  }
});

export async function protectSheet() {
  try {
    await Excel.run(async (context) => {
      let activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("protection/protected");

      await context.sync();

      if (!activeSheet.protection.protected) {
        activeSheet.protection.protect({ allowAutoFilter: true }, "123");
      }
    });
  } catch (error) {
    document.getElementById("run").style.backgroundColor = "#00FF00";
    document.getElementById("run").innerHTML = error;
  }
}

let externalDoc;

export async function getBase64() {
  // Retrieve the file and set up an HTML FileReader element.
  // const myFile = <HTMLInputElement>document.getElementById("file");
  const myFile = this.files[0];
  const reader = new FileReader();

  reader.onload = (event) => {
    // Remove the metadata before the base64-encoded string.
    const startIndex = reader.result.toString().indexOf("base64,");
    externalDoc = reader.result.toString().substr(startIndex + 7);
  };

  // Read the file as a data URL so that we can parse the base64-encoded string.
  if (myFile) {
    reader.readAsDataURL(myFile);
  }
}

export async function insertDocument() {
  try {
    await Word.run(async (context) => {
      // Retrieve the source document.
      context.document.getSelection().insertFileFromBase64(externalDoc, "Start", false);
      await context.sync();
    });
  } catch (error) {
    document.getElementById("run").style.backgroundColor = "#00FF00";
    document.getElementById("run").innerHTML = error;
  }
}

/*export function launchProjectInVSCode()
{
  const execSync = child_process.execSync;
  const command = `code D:\\testProjectCreation\\Word-Add-in-AI-Assistant`;
  try {
    //execSync(command, { timeout: 10000, stdio: ['ignore']});
    execSync(command);
    //vscode.window.showInformationMessage('Project launched successfully!');
  } catch (error) {
    //vscode.window.showErrorMessage('Failed to launch the project. ' + error);
    //console.log(error);
  }
}*/


export async function insertContentControl() {
  try {

      await Word.run(async (context) => {
      // Retrieve the source document.
      //context.document.getSelection().insertContentControl(Word.ContentControlType.plainText);
      const par = context.document.body.paragraphs.getFirst();
      par.load();
      await context.sync();
      var text = par.text;
      console.log(text);
    });
  } catch (error) {
    document.getElementById("run").style.backgroundColor = "#00FF00";
    document.getElementById("run").innerHTML = error;
  }
}

var dialog;

export async function save() {
  try {
    await Word.run(async (context) => {
      // Retrieve the source document.
      //var base64 = "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxMDAgMTAwIj4KICA8cGF0aCBkPSJNMzAsMWg0MGwyOSwyOXY0MGwtMjksMjloLTQwbC0yOS0yOXYtNDB6IiBzdHJva2U9IiMwMDAiIGZpbGw9Im5vbmUiLz4gCiAgPHBhdGggZD0iTTMxLDNoMzhsMjgsMjh2MzhsLTI4LDI4aC0zOGwtMjgtMjh2LTM4eiIgZmlsbD0iI2EyMyIvPiAKICA8dGV4dCB4PSI1MCIgeT0iNjgiIGZvbnQtc2l6ZT0iNDgiIGZpbGw9IiNGRkYiIHRleHQtYW5jaG9yPSJtaWRkbGUiPjwhW0NEQVRBWzQxMF1dPjwvdGV4dD4KPC9zdmc+Cg==";
      //context.document.body.insertInlinePictureFromBase64(base64, "Start");
      Office.context.ui.displayDialogAsync("https://app.gptzero.me/login", { height:50, width:50 }, dialogCallback);
      await context.sync();
    });
  } catch (error) {
    document.getElementById("run").style.backgroundColor = "#00FF00";
    document.getElementById("run").innerHTML = error;
  }
}

function dialogCallback(asyncResult) {
  if (asyncResult.status == "failed") {

      // In addition to general system errors, there are 3 specific errors for 
      // displayDialogAsync that you can handle individually.
      switch (asyncResult.error.code) {
          case 12004:
              console.log("Domain is not trusted");
              break;
          case 12005:
              console.log("HTTPS is required");
              break;
          case 12007:
              console.log("A dialog is already opened.");
              break;
          default:
              console.log(asyncResult.error.message);
              break;
      }
  }
  else {
      dialog = asyncResult.value;
      /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

      /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
      //dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
  }
}

function messageHandler(arg) {
  dialog.close();
  //const messageFromDialog = JSON.parse(arg.message);
  //console.log(messageFromDialog.name);
  console.log(arg.message);

  window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}


export async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

export async function run() {
  var element =  document.getElementById("scriptText") as HTMLTextAreaElement;
  var script = element.value;  
  console.log(script);
  var runFunc = new Function(script);
  runFunc();
}
