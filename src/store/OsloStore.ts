import Vuex from "vuex";
import Vue from "vue";
import { error, trace } from "../utils/Utils";
import { AppConfig } from "../utils/AppConfig";
import { IOsloItem } from "../oslo/IOsloItem";
import { Store } from "vuex";

Vue.use(Vuex);

export class OsloStore {
  private static instance: OsloStore;
  private store: any;

  private constructor() {
    this.init();
  }

  public static getInstance(): OsloStore {
    if (!OsloStore.instance) {
      OsloStore.instance = new OsloStore();
    }

    return OsloStore.instance;
  }

  // Fetches all the data from the Oslo database
  public init() {
    this.initializeStore();

    // only need to init once
    if (this.store.state.items.length < 1) {
      trace("initializing store");

      this.httpRequest("GET", AppConfig.dataFileUrl)
        .then((json: string) => {
          if (!json) {
            error("Oslo data empty");
          }
          const data = JSON.parse(json); //convert to usable JSON
          const cleandata = data["hits"]["hits"]; //filter out stuff we don't really need

          cleandata.map((item) => this.storeItem(item));

          trace("information stored in Vuex store");
        })
        .catch((error) => {
          trace("Error: " + error);
        });
    } else {
      trace("store already initialized");
    }
  }

  //Function to retrieve the data from an url
  private async httpRequest(verb: "GET" | "PUT", url: string): Promise<string> {
    return new Promise<string>((resolve, reject) => {
      const request = new XMLHttpRequest();

      // Callback after request.send()
      request.onload = function (event) {
        if (request.status === 200) {
          // HTTP request successful, resolve the promise with the response body
          resolve(request.response);
        } else {
          // HTTP request failed
          error(`Error after ${verb} from ${url} : ${request.status} ${request.statusText}`);
          resolve(null);
        }
      };

      request.open(verb, url, true /* async */);
      request.send();
    });
  }

  // Function to search the keyword in the Vuex store
  public osloStoreLookup(phrase: string, useExactMatching: boolean): IOsloItem[] {
    if (!phrase) {
      return null;
    }
    //clean
    phrase = phrase.toLowerCase().trim();
    // new list
    const matches: IOsloItem[] = [];

    let items = this.store.state.items;
    // loop for possible matches
    for (const item of items) {
      if (typeof item.label === "string") {
        //FIXME 4 objects are incomplete so we filter them out
        let possible = item.label.toLowerCase();
        let result = possible.search(phrase); // returns position of word in the label
        if (result >= 0) {
          // -1 is no match, so everything on position 0 to infinity is a match
          matches.push(item);
        }
      }
    }
    return matches.sort();
  }

  private storeItem(item) {
    let osloEntry: IOsloItem = {
      // new oslo object
      label: item["_source"]["prefLabel"],
      keyphrase: item["_source"]["id"],
      description: item["_source"]["definition"],
      reference: item["_source"]["context"],
    };
    this.store.commit("addItem", osloEntry);
  }

  private initializeStore() {
    this.store = new Store({
      state: {
        items: [] as IOsloItem[],
      },
      mutations: {
        addItem(state, item) {
          state.items.push(item);
        },
      },
    });
  }

  public getStore() {
    return this.store;
  }
}
