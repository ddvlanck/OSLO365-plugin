/*
 * Copyright (c) 2020 Vlaamse Overheid. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

export {};

import get = Reflect.get;

/** The config for this code */
namespace AppConfig {
  /** The URL of the Oslo data file. Retrieved with a simple GET. This static resource is replaced with a dynamic backend endpoint in the production environment. */
  export const dataFileUrl = "/oslo_terminology.json";

  //export const englishLocale = "/en.json";
  export const dutchLocale = "/nl.json";

  /** Set true to enable some trace messages to help debugging. */
  export const trace = true;
}

/** Delimiters used when splitting text into individual words */

/** After a word search, contains the list of matches with the same key */

/** Logs debug traces to the console, if enabled. */
function trace(text: string) {
  if (AppConfig.trace) {
    console.log(text);
  }
}

let excludedWords;

/** Logs error messages to the console. */
function error(text: string) {
  console.error(text);
}

function getDeleteButtons(): HTMLImageElement[] {
  const deletbtns: HTMLImageElement[] = [];

  const elements = document.getElementById("ExcludedWords").children;

  if (!elements) {
    return [];
  }

  for (let i = 0; i < elements.length; i++) {
    let element = elements[i];

    if (element instanceof HTMLImageElement && element.classList.contains("dictionary-deletebtn")) {
      deletbtns.push(element);
    }
  }

  return deletbtns;
}

function loadExcludedWords() {
  document.getElementById("ExcludedWords").innerHTML = "";

  // sort dictionary
  excludedWords.sort(function (a, b) {
    var textA = a[0].toUpperCase();
    var textB = b[0].toUpperCase();
    return textA < textB ? -1 : textA > textB ? 1 : 0;
  });

  for (let i = 0; i < excludedWords.length; i++) {
    document.getElementById("ExcludedWords").innerHTML += "<h3>" + excludedWords[i] + "</h3>";
    document.getElementById("ExcludedWords").innerHTML +=
      "<img class='dictionary-deletebtn' id='dictionary-deletebtn' style='float:left; padding: 8px 0px;' src='assets/deleteBtn.png' width='16px' data-id='" +
      i +
      "'>";
    document.getElementById("ExcludedWords").innerHTML += "<div class='clear'></div>";
  }

  for (const deleteBtn of getDeleteButtons()) {
    deleteBtn.onclick = function () {
      excludedWords.splice(deleteBtn.getAttribute("data-id"), 1);

      localStorage.setItem("excludedWords", JSON.stringify(excludedWords));
      loadExcludedWords();
    };
  }
}

/** Office calls this onReady handler to initialize the plugin */
Office.onReady((info) => {
  // This add-in is intended to be loaded in Word (2016 Desktop or Online)
  if (info.host === Office.HostType.Word) {
    // Initialize element visibility, register event handlers
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // There's a bug in the office.js libraries, causing an exception in some browsers when 3rd party cookies are blocked. Notify the user why the plugin doesn't work.
    try {
      let s = window.sessionStorage;
    } catch (error) {
      //setResultText("De extensie kon niet correct worden geladen. Gelieve in de browser instellingen alle cookies toe te laten, en daarna de pagina te herladen.");
    }

    if (localStorage.getItem("excludedWords") === null)
      excludedWords = ["ik", "jij", "hij", "zij", "wij", "jullie", "een", "de", "het"];
    else excludedWords = JSON.parse(localStorage.getItem("excludedWords"));

    loadExcludedWords();

    document.getElementById("add-excludedWord").onclick = addExcludedWord;
    document.getElementById("reset-plugin").onclick = resetPlugin;
  }

  function resetPlugin() {
    const btn = document.getElementById("reset-plugin");

    console.log(btn.innerHTML);

    if (btn.innerHTML != "ZEKER? (Klik om te bevestigen)") {
      btn.innerHTML = "ZEKER? (Klik om te bevestigen)";
      btn.style.backgroundColor = "red";
      btn.style.borderColor = "red";
    } else {
      localStorage.clear();
      location.reload();
    }
  }

  function addExcludedWord() {
    const wordToAdd = <HTMLInputElement>document.getElementById("newWordToExclude");

    if (wordToAdd.value.length >= 2) {
      if (excludedWords.indexOf(wordToAdd.value) == -1) {
        excludedWords.push(wordToAdd.value);
        localStorage.setItem("excludedWords", JSON.stringify(excludedWords));
        wordToAdd.value = "";
        document.getElementById("error").innerHTML = "";
      } else {
        document.getElementById("error").innerHTML = "Dit woord staat reeds in lijst.";
      }
    } else {
      document.getElementById("error").innerHTML = "Het woord moet langer zijn dan 2 tekens.";
    }

    loadExcludedWords();
  }
});
