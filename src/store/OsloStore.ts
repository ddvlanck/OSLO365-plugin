import Vuex from 'vuex';
import Vue from 'vue';
import {error, trace} from "../utils/Utils";
import {AppConfig} from "../utils/AppConfig";


Vue.use(Vuex);
getData();

//Vuex store
export const store = new Vuex.Store({
    state: {
        title: "bliepblop",
        definition: "blopblop",
        url: "www.blopblop.com"
    },
    mutations: {
        updateTitle (state, title) {
            state.title = title
        },
        updateDefinition (state, definition) {
            state.definition = definition
        },
        updateUrl (state, url) {
            state.url = url
        }
    }
});

//fetches all the data from the Oslo database
function getData(){
    httpRequest("GET", AppConfig.dataFileUrl).then((json: string) => {
        if (!json) {
            error('Oslo data empty');
        }
        //clean on the objects we need
        const data = JSON.parse(json);
        let title = data["hits"]["hits"][5]["_source"]["prefLabel"];  //5 to grab only 1 object - test purposes
        let definition = data["hits"]["hits"][5]["_source"]["definition"];
        let url = data["hits"]["hits"][5]["_source"]["id"];

        store.commit('updateTitle', title)
        store.commit('updateDefinition', definition)
        store.commit('updateUrl', url)

        console.log(store.state.definition);
        console.log(store.state.title);
        console.log(store.state.url);

    }).catch((error) => {
        trace("Error: " + error);
    });
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