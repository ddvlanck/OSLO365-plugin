/*
 * Copyright (c) 2020 Vlaamse Overheid. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

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

let myDictionary;

/** Logs error messages to the console. */
function error(text: string) {
  console.error(text);
}
 
function createSearchResultItemHtml(
  description: string,
  referenceUrl: string
): string {
  let html = "";

  description = escapeHtml(description);
  let referenceUrlEscaped = escapeHtml(referenceUrl);

  html += `
	  <p>
		${description}<p>
		<a href="${referenceUrl}" target="_blank">${referenceUrlEscaped}</a><p>
		<hr>`;

  return html;
}

function escapeHtml(text: string) {
  return text ? text.replace(/[^0-9A-Za-z ]/g, (char) => "&#" + char.charCodeAt(0) + ";") : "";
}

function getDeleteButtons(): HTMLImageElement[] {
  const deletbtns: HTMLImageElement[] = [];

  const elements = document.getElementById("ResultBox").children;

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

function loadDictionary()
{
  document.getElementById("ResultBox").innerHTML = "";

  // sort dictionary
  myDictionary.sort(function(a, b) {
    var textA = a[0].toUpperCase();
    var textB = b[0].toUpperCase();
    return (textA < textB) ? -1 : (textA > textB) ? 1 : 0;
});

  for(let i = 0; i < myDictionary.length; i++)
    {
    const deleteBtn = "<img class='dictionary-deletebtn' id='dictionary-deletebtn' style='float:left;' src='assets/deleteBtn.png' width='16px' data-id='" + i + "'>";

      document.getElementById("ResultBox").innerHTML += "<h3 style='float:left;margin-block-start: 0px;margin-block-end: 0px;margin-right: 9px;'>" + myDictionary[i][0].toLowerCase() + "</h3>";
      document.getElementById("ResultBox").innerHTML += '<i class="vi vi-trash" aria-hidden="true"></i>';
      document.getElementById("ResultBox").innerHTML += deleteBtn;
      document.getElementById("ResultBox").innerHTML += "<div class='clear'></div>";
      document.getElementById("ResultBox").innerHTML += createSearchResultItemHtml(myDictionary[i][1].description, myDictionary[i][1].reference);
    }

    for (const deleteBtn of getDeleteButtons()) {
      deleteBtn.onclick = function(){
        //let wordsToIgnore = JSON.parse(localStorage.getItem('wordsToIgnore'));
        //let index = wordsToIgnore.indexOf(myDictionary[deleteBtn.getAttribute("data-id")][0]);
        myDictionary.splice(deleteBtn.getAttribute("data-id"),1);

        

        
        //wordsToIgnore.splice(index,1)

        //localStorage.setItem('wordsToIgnore', JSON.stringify(wordsToIgnore));

        localStorage.setItem('myDictionary', JSON.stringify(myDictionary));        
        loadDictionary();
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

    if(localStorage.getItem("myDictionary") === null)
    myDictionary = [];
    else
    myDictionary = JSON.parse(localStorage.getItem('myDictionary'));

    
    
    loadDictionary();

    




  }


});

