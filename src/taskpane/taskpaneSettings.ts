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

    




  }


});

