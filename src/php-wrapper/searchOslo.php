<?php
/*
 * Copyright (c) 2020 Vlaamse Overheid. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// This is a simple HTTP(s) handler that acts as a proxy/wrapper around the Elasticsearch API for the OSLO test database.
// This endpoint can be installed on an Apache with support for PHP 5.5+.
// By serving the plugin from the same web server, this avoids issues due to same-origin and mixed-content restrictions in the browsers.
// The endpoint pattern is simply: {BASE_URL}/searchOslo?q={search_phrase}

class OsloElasticsearchApi
{
	public function search($phrase = "") {
		$searchUrl = "http://52.143.54.172:9200/terminology/_search";

		$newline = "\r\n";
		
		$httpOptions = array();
		
		// Basic request configuration
		$httpOptions["method"] = "POST";
		$httpOptions["timeout"] = 10.0; // Actual timeout is about 20s (x2)
		$httpOptions["ignore_errors"] = true;
		
		// Payload
		$data = json_decode('{ "query": { "wildcard": { "prefLabel": { "value": "*' . $phrase . '*", "boost": 1.0, "rewrite": "constant_score" } } } }');
		$data->query->wildcard->prefLabel->value = "*$phrase*";
		$content = json_encode($data);

		$httpOptions["content"] = $content;
		
		// Header
		$header = "";
		$header .= "Accept: application/json" . $newline;
		$header .= "Content-Type: application/json" . $newline;
		$header .= "Content-Length: " . strlen($content) . $newline;
		
		$httpOptions['header'] = $header . $newline;

		// Do the HTTP call
		$options = array('http' => $httpOptions);

		$context = stream_context_create($options);
		
		$json = file_get_contents($searchUrl, false, $context);

		return $json ? $json : null;
	}
}

header("Content-Type: application/json; charset=utf-8");
// Make sure the backend call doesn't get blocked because of CORS
header("Access-Control-Allow-Origin: *");

$phrase = $_REQUEST["q"];

$api = new OsloElasticsearchApi();
echo $api->search($phrase ? $phrase : "");
?>