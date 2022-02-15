import Vuex from 'vuex';
import Vue from 'vue';
import {error, trace} from "../utils/Utils";
import {AppConfig} from "../utils/AppConfig";


Vue.use(Vuex);
getdata();

export const store = new Vuex.Store({
    state: {test: "bliepblop"}
});

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
function getdata(){
    // The first cache is a simple list of Oslo result items
    // Load the data from the web server. We're assuming a simple GET without authentication.
    httpRequest("GET", AppConfig.dataFileUrl).then((json: string) => {
        if (!json) {
            error('Oslo data empty');
        }

        const data = JSON.parse(json);
        let results = data["hits"]["hits"];
        console.log(results);

    }).catch((error) => {
        trace("Error: " + error);
    });
}