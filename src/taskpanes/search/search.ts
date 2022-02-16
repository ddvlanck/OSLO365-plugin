import Vue from "vue";
import Vuex from "vuex";
import root from "./pages/Root.vue";
const VlUiVueComponents = require("@govflanders/vl-ui-vue-components");
import { trace } from "../../utils/Utils";
import { OsloCache } from "../../oslo/OsloCache";
import EventBus from "../../utils/EventBus";
import {getData} from "../../store/OsloStore";

let searching = false;

// configuration of the built-in validator
const validatorConfig = {
  inject: true,
  locale: "nl",
};

Vue.use(VlUiVueComponents, {
  validation: validatorConfig,
});
Vue.use(Vuex);

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    var app = new Vue({
      el: "#app",
      render: (h) => h(root),
    });
  }

  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onWordSelectionChanged);
  getData();
  OsloCache.init();
});

/** Called when the user selects something in the Word document */
function onWordSelectionChanged(result: Office.AsyncResult<void>) {
  processSelection();
}

/** Uses the current selection to perform a search in the OSLO data set. */
function processSelection() {
  // Callback after reading selected text
  let onDataSelected = function (asyncResult) {
    let error = asyncResult.error;

    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      error("Selection failed: " + error.name + "; " + error.message);
    } else {
      // The selected text is used as a search phrase
      let searchPhrase = asyncResult.value ? asyncResult.value.trim() : "";
      EventBus.$emit("onWordSelection", searchPhrase);

      if (searching) {
        // When using the "Volgende Zoeken" button, enforce exact matching
        searchPhrase = searchPhrase ? "=" + searchPhrase : "";
        searching = false;
      }
      trace("processSelection [" + searchPhrase + "]");
      search(searchPhrase);
    }
  };

  // Get the currently selected text from the Word document, and process it
  Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text,
    { valueFormat: "unformatted", filterType: "all" },
    onDataSelected
  );
}

/** Searches a given phrase in the OSLO data set. */
export function search(searchPhrase: string) {
  console.log(`Looking for "${searchPhrase}"`);

  if (!searchPhrase) {
    return;
  }

  // If the search phrase begins with an equals char, perform an exact match (otherwise a "contains" match)
  const exactMatch = searchPhrase.charAt(0) == "=";

  if (exactMatch) {
    // Remove the equals char from the search phrase
    searchPhrase = searchPhrase.slice(1);
  }

  const osloCache = OsloCache.getInstance();

  // Search the phrase in the OSLO database
  const osloResult = osloCache.osloLookup(searchPhrase, exactMatch);

  EventBus.$emit("onSearchResult", osloResult);
}
