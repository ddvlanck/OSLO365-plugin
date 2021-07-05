/** The port node.js starts a web server on, always on https://127.0.0.1 . */
const dev_webserver_port = process.env.npm_package_config_dev_server_port || 3000;

/** The host URL and domain used to generate the /dist/manifest.xml */
const plugin_host_url = process.env.npm_package_config_plugin_host_url || "https://oslo.mywebserver.dev/";
const plugin_host_domain = process.env.npm_package_config_plugin_host_domain || "oslo.mywebserver.dev";


const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const fs = require("fs");
const webpack = require("webpack");

module.exports = async (env, options) => {
	const dev = options.mode === "development";
	const config = {
		devtool: "source-map",
		entry: {
			polyfill: "@babel/polyfill",
			taskpane: "./src/taskpane/taskpane.ts",
			taskpanetwo: "./src/taskpane/taskpanetwo.ts"
		},
		resolve: {
			extensions: [".ts", ".tsx", ".html", ".js"]
		},
		module: {
			rules: [{
				test: /\.ts$/,
				exclude: /node_modules/,
				use: "babel-loader"
			}, {
				test: /\.tsx?$/,
				exclude: /node_modules/,
				use: "ts-loader"
			}, {
				test: /\.html$/,
				exclude: /node_modules/,
				use: "html-loader"
			}, {
				test: /\.(png|jpg|jpeg|gif|json)$/,
				use: "file-loader"
			}]
		},
		plugins: [
			new CleanWebpackPlugin(),

			new HtmlWebpackPlugin({
				filename: "taskpane.html",
				template: "./src/taskpane/taskpane.html",
				chunks: ["polyfill", "taskpane"]
			}),

			new HtmlWebpackPlugin({
				filename: "taskpanetwo.html",
				template: "./src/taskpane/taskpanetwo.html",
				chunks: ["polyfill", "taskpanetwo"]
			}),

			new CopyWebpackPlugin([ {
				from: "./src/taskpane/taskpane.css",
				to: "taskpane.css"
			}, {
				from: "./assets/vo_logo_32.png",
				to: "assets/vo_logo_32.png"
			},{
				from: "./assets/vo_logo_64.png",
				to: "assets/vo_logo_64.png"
			}, {
				from: "./assets/vo_oslo_logo.png",
				to: "assets/vo_oslo_logo.png"
			}, {
				from: "./assets/vo_logo_large.png",
				to: "assets/vo_logo_large.png"
			}, {
				from: "./src/taskpane/oslo_terminology.json",
				to: "oslo_terminology.json"
			}, {
				from: "./manifest.template.xml",
				to: "manifest.xml",
				transform: (content, path) => replace_placeholders_dist(content)
			}, {
				from: "./manifest.template.xml",
				to: "../manifest.xml",
				force: true,
				transform: (content, path) => replace_placeholders_dev(content)
			}, {
				from: "./src/locales/en.json",
				to: "en.json"
			}, {
				from: "./src/locales/nl.json",
				to: "nl.json"
			}
		])
	],
	devServer: {
	headers: {
		"Access-Control-Allow-Origin": "*"
	},
	https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
	port: dev_webserver_port
	}
};

return config;
};

const template_generated_from = "Generated from manifest.template.xml";

function replace_placeholders_dist(content) {
	return content.toString()
		.replace(/{{template_info}}/, template_generated_from + " (deploy mode)")
		.replace(/{{plugin_host_url}}/g, plugin_host_url)
		.replace(/{{plugin_host_domain}}/g, plugin_host_domain);
}

function replace_placeholders_dev(content) {
	return content.toString()
		.replace(/{{template_info}}/, template_generated_from + " (local test mode)")
		.replace(/{{plugin_host_url}}/g, "https://127.0.0.1:" + dev_webserver_port + "/")
		.replace(/{{plugin_host_domain}}/g, "localhost");
}
