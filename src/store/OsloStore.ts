import Vuex from 'vuex';
import Vue from 'vue';
import {error, trace} from "../utils/Utils";
import {AppConfig} from "../utils/AppConfig";


Vue.use(Vuex);
getData();

//Vuex store
export const store = new Vuex.Store({
    state: {
        title: [],
        definition: [],
        url: []
    },
    mutations: {
        addTitle (state, title) {
            state.title.push(title);
        },
        addDefinition (state, definition) {
            state.definition.push(definition);
        },
        addUrl (state, url) {
            state.url.push(url);
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

        //add all titles
        for (let i = 0; i < data["hits"]["hits"].length; i++) {
            let title = data["hits"]["hits"][i]["_source"]["prefLabel"];
            store.commit('addTitle', title);
        }
        // add all definitions
        for (let i = 0; i < data["hits"]["hits"].length; i++) {
            let definition = data["hits"]["hits"][i]["_source"]["definition"];
            store.commit('addDefinition', definition);
        }
        // add all url's
        for (let i = 0; i < data["hits"]["hits"].length; i++) {
            let url = data["hits"]["hits"][5]["_source"]["id"];
            store.commit('addUrl', url);
        }

        console.log(store.state.title);
        console.log(store.state.definition);
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