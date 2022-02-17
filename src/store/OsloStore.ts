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

//fetches all the data from the Oslo database
export function initStore(){
    // only need to init once
    if (store.state.items.length < 1){
        trace("initializing store");

        httpRequest("GET", AppConfig.dataFileUrl).then((json: string) => {
            if (!json) {
                error('Oslo data empty');
            }
            //convert to usable JSON
            const data = JSON.parse(json);

            // saves all items as OsloItems (interface) in Vuex store
            for (let i = 0; i < data["hits"]["hits"].length; i++) {

                let label = data["hits"]["hits"][i]["_source"]["prefLabel"];
                let id = data["hits"]["hits"][i]["_source"]["id"];
                let definition = data["hits"]["hits"][i]["_source"]["definition"];
                let context = data["hits"]["hits"][i]["_source"]["context"];

                let osloEntry: IOsloItem = {
                    label: label,
                    keyphrase: id,
                    description: definition,
                    reference: context
                };
                store.commit('addItem', osloEntry);
            }
            trace("information stored in Vuex store");

        }).catch((error) => {
            trace("Error: " + error);
        });
    }
    else {
        trace("store already initialized");
    }
}
//Function to retrieve the data from an url
async function httpRequest(verb: "GET" | "PUT", url: string): Promise<string> {
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
export function osloStoreLookup(phrase: string, useExactMatching: boolean): IOsloItem[] {
    if (!phrase) {
        return null;
    }
    //clean
    phrase = phrase.toLowerCase().trim();
    // new list
    const matches: IOsloItem[] = [];

    let items = store.state.items;
    // loop for possible matches
    for (const item of items){
        if (typeof item.label === 'string'){
            let possible = item.label.toLowerCase();
            let result = possible.search(phrase); // returns position of word in the label
            if (result >= 0){  // -1 is no match, so everything on position 0 to infinity is a match
                matches.push(item);
            }
        }
    }
    matches.sort();
    return matches

}
