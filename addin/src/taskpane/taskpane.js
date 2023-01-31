/*
  Copyright (c) 2023 Stefano Aldegheri

  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all
  copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  SOFTWARE.
*/

/* global document, Office */

import { globalOptions } from "./GlobalOptions";
var initGui = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    var tabcontent = document.getElementsByClassName("tabcontent");
    for (var i = 0; i < tabcontent.length; i++) {
      tabcontent[i].style.display = "none";
    }
    document.getElementById("run").onclick = run;
    document.getElementById("txtAPIKEY").onblur = valuesChanged;
    document.getElementById("cmbLanguage").onchange = valuesChanged;
    document.getElementById("cmbGPTModel").onchange = valuesChanged;
    document.getElementById("chkSetMaxNumWords").onchange = valuesChanged;
    document.getElementById("iiMaxNumWords").onblur = valuesChanged;
    document.getElementById("diTemperature").onblur = valuesChanged;
    document.getElementById("txtFormat").onblur = valuesChanged;
    document.getElementById("btnSaveOptions").onclick = saveOptions;
    document.getElementById("btnSubstituteSelectionInMail").onclick = substituteSelectionInMail;

    document.getElementById("tbBtnRephrase").onclick = showTabRephraser;
    document.getElementById("tbBtnOptions").onclick = showTabOptions;
    loadOptions();
    showTabRephraser();
    document.getElementById("btnSubstituteSelectionInMail").style.visibility = "hidden";
  }
});

export function saveOptions() {
  try {
    Office.context.roamingSettings.set("gptKey", globalOptions.gptKey);
    Office.context.roamingSettings.set("format", globalOptions.format);
    Office.context.roamingSettings.set("language", globalOptions.language);
    Office.context.roamingSettings.set("gptModel", globalOptions.gptModel);
    Office.context.roamingSettings.set("setMaxNumWords", globalOptions.setMaxNumWords);
    Office.context.roamingSettings.set("maxNumWords", globalOptions.maxNumWords);
    Office.context.roamingSettings.set("temperature", globalOptions.temperature);
    Office.context.roamingSettings.saveAsync();
  } catch (error) {}
}

export function loadOptions() {
  try {
    const gptKey = Office.context.roamingSettings.get("gptKey");
    if (typeof gptKey === typeof globalOptions.gptKey) globalOptions.gptKey = gptKey;

    const format = Office.context.roamingSettings.get("format");
    if (typeof format === typeof globalOptions.format) globalOptions.format = format;

    const language = Office.context.roamingSettings.get("language");
    if (typeof language === typeof globalOptions.language) globalOptions.language = language;

    const gptModel = Office.context.roamingSettings.get("gptModel");
    if (typeof gptModel === typeof globalOptions.gptModel) globalOptions.gptModel = gptModel;

    const setMaxNumWords = Office.context.roamingSettings.get("setMaxNumWords");
    if (typeof setMaxNumWords === typeof globalOptions.setMaxNumWords) globalOptions.setMaxNumWords = setMaxNumWords;

    const maxNumWords = Office.context.roamingSettings.get("maxNumWords");
    if (typeof maxNumWords === typeof globalOptions.maxNumWords) globalOptions.maxNumWords = maxNumWords;

    const temperature = Office.context.roamingSettings.get("temperature");
    if (typeof temperature === typeof globalOptions.temperature) globalOptions.temperature = temperature;
  } catch (error) {}
}

function dataToGUI() {
  document.getElementById("txtAPIKEY").value = globalOptions.gptKey;
  document.getElementById("txtFormat").value = globalOptions.format;
  document.getElementById("cmbLanguage").value = globalOptions.language;
  document.getElementById("cmbGPTModel").value = globalOptions.gptModel;
  document.getElementById("chkSetMaxNumWords").checked = globalOptions.setMaxNumWords;
  document.getElementById("iiMaxNumWords").value = globalOptions.maxNumWords;
  document.getElementById("diTemperature").value = globalOptions.temperature;
  toggleEnabled("chkSetMaxNumWords", "iiMaxNumWords");
}

function toggleEnabled(checkboxID, controlID, reverse = false) {
  var chk1 = document.getElementById(checkboxID);
  var txt1 = document.getElementById(controlID);
  if (chk1.checked) {
    txt1.disabled = reverse;
  } else {
    txt1.disabled = !reverse;
  }
}

function guiToData() {
  globalOptions.gptKey = document.getElementById("txtAPIKEY").value;
  globalOptions.format = document.getElementById("txtFormat").value;
  globalOptions.language = document.getElementById("cmbLanguage").value;
  globalOptions.gptModel = document.getElementById("cmbGPTModel").value;
  globalOptions.setMaxNumWords = document.getElementById("chkSetMaxNumWords").checked;
  globalOptions.maxNumWords = Math.max(1, Math.min(2040, document.getElementById("iiMaxNumWords").value));
  globalOptions.temperature = Math.max(0.0, Math.min(1.0, document.getElementById("diTemperature").value));
}

function valuesChanged() {
  if (initGui) return;

  initGui = true;
  guiToData();
  dataToGUI();
  initGui = false;
}

function showTabRephraser() {
  showTab("tbLblRephrase", "Rephraser");
}

function showTabOptions() {
  showTab("tbLblOptions", "Options");
}

function showTab(tabLabel, tabName) {
  var i, tabcontent;
  tabcontent = document.getElementsByClassName("tabcontent");
  for (i = 0; i < tabcontent.length; i++) {
    tabcontent[i].style.display = "none";
  }

  var tbLblRephrase = document.getElementById("tbLblRephrase");
  var tbLblOptions = document.getElementById("tbLblOptions");
  tbLblRephrase.className = tbLblRephrase.className.replace(" lblChecked", "");
  tbLblOptions.className = tbLblOptions.className.replace(" lblChecked", "");

  dataToGUI();
  document.getElementById(tabName).style.display = "flex";
  document.getElementById(tabLabel).className += " lblChecked";
}

function modelToText(model) {
  switch (model) {
    case "ada":
      return "text-ada-001";
    case "babbage":
      return "text-babbage-001";
    case "curie":
      return "text-curie-001";
    case "davinci":
      return "text-davinci-003";
    default:
      return "text-curie-001";
  }
}

async function makeGPTRequest(prompt, api_key, max_tokens, modelStr, temperature) {
  const MODEL = modelToText(modelStr);
  const xhr = new XMLHttpRequest();
  xhr.open("POST", "https://api.openai.com/v1/completions");
  xhr.setRequestHeader("Content-Type", "application/json");
  xhr.setRequestHeader("Authorization", `Bearer ${api_key}`);

  return new Promise((resolve, reject) => {
    xhr.onload = () => {
      try {
        if (xhr.status === 200) {
          resolve(JSON.parse(xhr.response));
        } else if (xhr.responseText != "") {
          reject(JSON.parse(xhr.responseText));
        } else {
          reject(xhr.statusText);
        }
      } catch (e) {
        reject(e.message);
      }
    };

    xhr.send(
      JSON.stringify({
        model: MODEL,
        prompt: prompt,
        max_tokens: max_tokens,
        temperature: temperature,
      })
    );
  });
}

async function rephraseText(prompt) {
  var nTok = 2040;
  if (globalOptions.model === "davinci") nTok = 4000;

  try {
    const resultText = await makeGPTRequest(
      prompt,
      globalOptions.gptKey,
      nTok,
      globalOptions.gptModel,
      globalOptions.temperature
    );
    if (typeof resultText === "string") {
      return ["Error:<br>" + resultText, false];
    } else if ("choices" in resultText) {
      return [resultText["choices"][0]["text"], true];
    } else if ("error" in resultText) {
      return ["Error:<br>" + resultText["error"]["message"], false];
    } else {
      return ["Error:<br>Generic Error", false];
    }
  } catch (ex) {
    if (typeof ex === "string") {
      return ["Error:<br>" + ex, false];
    } else if ("choices" in ex) {
      return [ex["choices"][0]["text"], true];
    } else if ("error" in ex) {
      return ["Error:<br>" + ex["error"]["message"], false];
    } else {
      return ["Error:<br>Generic Error", false];
    }
  }
}

export async function run() {
  var PROD = true;
  const NL = "<br/>";
  document.getElementById("btnSubstituteSelectionInMail").style.visibility = "hidden";
  document.getElementById("lblRephrasedText").style.visibility = "hidden";
  document.getElementById("lblRephrasedText").className = "";

  Office.context.mailbox.item.getSelectedDataAsync(
    Office.CoercionType.Text,
    { valueFormat: "unformatted" },
    async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        var prompt = "";
        if (globalOptions.format != "") {
          var bodyData = result.value.data;

          var bodyDataStr = "";
          if (typeof bodyData === "string") {
            bodyDataStr = bodyData;
          } else {
            bodyDataStr = String(bodyData);
          }

          if (bodyData != "") {
            prompt = "If necessary translate in " + globalOptions.language;

            if (globalOptions.setMaxNumWords) {
              prompt += " and summarize in maximum " + globalOptions.maxNumWords + " words ";
            } else {
              prompt += " and paraphrase ";
            }

            prompt += "like a " + globalOptions.format + " without adding any additional information: " + bodyDataStr;
          }

          if (prompt != "") {
            var finalText = "";
            if (PROD) {
              const resultText = await rephraseText(prompt);
              finalText = resultText[0];
              if (resultText[1]) document.getElementById("btnSubstituteSelectionInMail").style.visibility = "visible";
            } else {
              finalText = prompt;
            }

            document.getElementById("lblRephrasedText").innerHTML = finalText
              .trimStart()
              .replaceAll("\r\n", NL)
              .replaceAll("\r", NL)
              .replaceAll("\n", NL);

            document.getElementById("lblRephrasedText").style.visibility = "visible";
          } else {
            document.getElementById("lblRephrasedText").style.visibility = "visible";
            document.getElementById("lblRephrasedText").innerHTML = "Error: No text selected";
          }
        } else {
          document.getElementById("lblRephrasedText").style.visibility = "visible";
          document.getElementById("lblRephrasedText").innerHTML = "Error: Invalid format selected";
        }
      } else {
        document.getElementById("lblRephrasedText").style.visibility = "visible";
        document.getElementById("lblRephrasedText").innerHTML = "Error: No text selected";
      }
      document.getElementById("lblRephrasedText").className = "lblWordWrap";
    }
  );
}

export async function substituteSelectionInMail() {
  var text = document.getElementById("lblRephrasedText").innerHTML;

  if (text !== "") {
    const formattedText = text.replaceAll("<br/>", "\n").trimStart();
    Office.context.mailbox.item.setSelectedDataAsync(formattedText);
  }
}
