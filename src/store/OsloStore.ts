import Vuex from 'vuex';
import Vue from 'vue';
import {error, trace} from "../utils/Utils";
import {AppConfig} from "../utils/AppConfig";
import {IOsloItem} from "../oslo/IOsloItem";

Vue.use(Vuex);

//Vuex store
export const store = new Vuex.Store({
    state: {
        items: []
    },
    mutations: {
        addItem (state, item) {
            state.items.push(item)
        }
    }
});
export class OsloStore {

    private static instance: OsloStore;

    private constructor() {
        this.initStore();
    }

    public static getInstance(): OsloStore {
        if (!OsloStore.instance) {
            OsloStore.instance = new OsloStore();
        }

        return OsloStore.instance;
    }
    //fetches all the data from the Oslo database
    public initStore() {
        // only need to init once
        if (store.state.items.length < 1) {
            trace("initializing store");

            this.httpRequest("GET", AppConfig.dataFileUrl).then((json: string) => {
                if (!json) {
                    error('Oslo data empty');
                }
                const data = JSON.parse(json); //convert to usable JSON
                const cleandata = data["hits"]["hits"]; //filter out stuff we don't really need

                cleandata.map(item => OsloStore.storeItem(item));

                trace("information stored in Vuex store");

            }).catch((error) => {
                trace("Error: " + error);
            });
        } else {
            trace("store already initialized");
        }
    }
    //Function to retrieve the data from an url
    async httpRequest(verb: "GET" | "PUT", url: string): Promise<string> {
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
            }

            request.open(verb, url, true /* async */);
            request.send();
        });
    }
    //function to search the keyword in the Vuex store
    public osloStoreLookup(phrase: string, useExactMatching: boolean): IOsloItem[] {
        if (!phrase) {
            return null;
        }
        //clean
        phrase = phrase.toLowerCase().trim();
        // new list
        const matches: IOsloItem[] = [];

        let items = store.state.items;
        // loop for possible matches
        for (const item of items) {
            if (typeof item.label === 'string') {
                let possible = item.label.toLowerCase();
                let result = possible.search(phrase); // returns position of word in the label
                if (result >= 0) {  // -1 is no match, so everything on position 0 to infinity is a match
                    matches.push(item);
                }
            }
        }
        return matches.sort();
    }

    private static storeItem(item) {
        let osloEntry: IOsloItem = { // new oslo object
            label: item["_source"]["prefLabel"],
            keyphrase: item["_source"]["id"],
            description: item["_source"]["definition"],
            reference: item["_source"]["context"]
        };
        store.commit('addItem', osloEntry);
    }
}