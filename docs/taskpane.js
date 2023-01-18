/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { globalOptions } from "./GlobalOptions";
var initGui = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
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

    var tablinks = document.getElementsByClassName("tablinks");
    tablinks[0].onclick = showTabRephraser;
    tablinks[1].onclick = showTabOptions;
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
  showTab(0, "Rephraser");
}

function showTabOptions() {
  showTab(1, "Options");
}

function showTab(tabButton, tabName) {
  var i, tabcontent, tablinks;
  tabcontent = document.getElementsByClassName("tabcontent");
  for (i = 0; i < tabcontent.length; i++) {
    tabcontent[i].style.display = "none";
  }
  tablinks = document.getElementsByClassName("tablinks");
  for (i = 0; i < tablinks.length; i++) {
    tablinks[i].className = tablinks[i].className.replace(" active", "");
  }
  dataToGUI();
  document.getElementById(tabName).style.display = "block";
  tablinks[tabButton].className += " active";
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
      if (xhr.status === 200) {
        resolve(JSON.parse(xhr.response));
      } else {
        reject(xhr.statusText);
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
  const nTok = 2040;

  const resultText = await makeGPTRequest(
    prompt,
    globalOptions.gptKey,
    nTok,
    globalOptions.gptModel,
    globalOptions.temperature
  );
  if (resultText["error"]) {
    return "Error:<br>" + resultText["error"]["message"];
  } else {
    return resultText["choices"][0]["text"];
  }
}

export async function run() {
  const USE_GPT = true;
  document.getElementById("btnSubstituteSelectionInMail").style.visibility = "hidden";

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
            if (USE_GPT) {
              const resultText = await rephraseText(prompt);
              document.getElementById("lblRephrasedText").innerHTML = resultText.trimStart();
            } else {
              document.getElementById("lblRephrasedText").innerHTML = "NO GPT IN USE:<br><br>" + prompt;
            }
            document.getElementById("btnSubstituteSelectionInMail").style.visibility = "visible";
          } else {
            document.getElementById("lblRephrasedText").innerHTML = "Error: No text selected";
          }
        } else {
          document.getElementById("lblRephrasedText").innerHTML = "Error: Invalid format selected";
        }
      } else {
        document.getElementById("lblRephrasedText").innerHTML = "Error: No text selected";
      }
    }
  );
}

export async function substituteSelectionInMail() {
  var text = document.getElementById("lblRephrasedText").innerHTML;

  if (text !== "") {
    const formattedText = text.replace("<br>", "\n").trimStart();
    Office.context.mailbox.item.setSelectedDataAsync(formattedText);
  }
}
