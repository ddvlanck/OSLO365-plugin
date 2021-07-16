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
			begrippenkader: "./src/taskpane/begrippenkader.ts",
			documentControle: "./src/taskpane/documentControle.ts",
			mijnWoordenboek: "./src/taskpane/mijnWoordenboek.ts",
			instellingen: "./src/taskpane/instellingen.ts",
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
				filename: "begrippenkader.html",
				template: "./src/taskpane/begrippenkader.html",
				chunks: ["polyfill", "begrippenkader"]
			}),

			new HtmlWebpackPlugin({
				filename: "documentControle.html",
				template: "./src/taskpane/documentControle.html",
				chunks: ["polyfill", "documentControle"]
			}),

			new HtmlWebpackPlugin({
				filename: "mijnWoordenboek.html",
				template: "./src/taskpane/mijnWoordenboek.html",
				chunks: ["polyfill", "mijnWoordenboek"]
			}),

			new HtmlWebpackPlugin({
				filename: "instellingen.html",
				template: "./src/taskpane/instellingen.html",
				chunks: ["polyfill", "instellingen"]
			}),

			new HtmlWebpackPlugin({
				filename: "functions.html",
				template: "./src/taskpane/functions.html",
			}),

			new HtmlWebpackPlugin({
				filename: "help.html",
				template: "./src/taskpane/help.html",
				chunks: ["polyfill"]
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
			},
			{
				from: "./assets/icons/80x80/begrippenkader.png",
				to: "assets/icons/80x80/begrippenkader.png"
			},
			{
				from: "./assets/icons/32x32/begrippenkader.png",
				to: "assets/icons/32x32/begrippenkader.png"
			},
			{
				from: "./assets/icons/16x16/begrippenkader.png",
				to: "assets/icons/16x16/begrippenkader.png"
			},
			{
				from: "./assets/icons/80x80/over.png",
				to: "assets/icons/80x80/over.png"
			},
			{
				from: "./assets/icons/32x32/over.png",
				to: "assets/icons/32x32/over.png"
			},
			{
				from: "./assets/icons/16x16/over.png",
				to: "assets/icons/16x16/over.png"
			},
			{
				from: "./assets/icons/80x80/help.png",
				to: "assets/icons/80x80/help.png"
			},
			{
				from: "./assets/icons/32x32/help.png",
				to: "assets/icons/32x32/help.png"
			},
			{
				from: "./assets/icons/16x16/help.png",
				to: "assets/icons/16x16/help.png"
			},
			{
				from: "./assets/loading.gif",
				to: "assets/loading.gif"
			},
			{
				from: "./assets/deleteBtn.png",
				to: "assets/deleteBtn.png"
			},
			{
				from: "./assets/icons/80x80/documentcontrole.png",
				to: "assets/icons/80x80/documentcontrole.png"
			},
			{
				from: "./assets/icons/32x32/documentcontrole.png",
				to: "assets/icons/32x32/documentcontrole.png"
			},
			{
				from: "./assets/icons/16x16/documentcontrole.png",
				to: "assets/icons/16x16/documentcontrole.png"
			},
			{
				from: "./assets/icons/80x80/mijnwoordenboek.png",
				to: "assets/icons/80x80/mijnwoordenboek.png"
			},
			{
				from: "./assets/icons/32x32/mijnwoordenboek.png",
				to: "assets/icons/32x32/mijnwoordenboek.png"
			},
			{
				from: "./assets/icons/16x16/mijnwoordenboek.png",
				to: "assets/icons/16x16/mijnwoordenboek.png"
			},
			{
				from: "./assets/icons/80x80/instellingen.png",
				to: "assets/icons/80x80/instellingen.png"
			},
			{
				from: "./assets/icons/32x32/instellingen.png",
				to: "assets/icons/32x32/instellingen.png"
			},
			{
				from: "./assets/icons/16x16/instellingen.png",
				to: "assets/icons/16x16/instellingen.png"
			},
			{
				from: "./assets/icons/80x80/proximus.png",
				to: "assets/icons/80x80/proximus.png"
			},
			{
				from: "./assets/icons/32x32/proximus.png",
				to: "assets/icons/32x32/proximus.png"
			},
			{
				from: "./assets/icons/16x16/proximus.png",
				to: "assets/icons/16x16/proximus.png"
			},
			{
				from: "./assets/icons/80x80/microsoft.png",
				to: "assets/icons/80x80/microsoft.png"
			},
			{
				from: "./assets/icons/32x32/microsoft.png",
				to: "assets/icons/32x32/microsoft.png"
			},
			{
				from: "./assets/icons/16x16/microsoft.png",
				to: "assets/icons/16x16/microsoft.png"
			},
			{
				from: "./src/taskpane/oslo_terminology.json",
				to: "oslo_terminology.json"
			},
			{
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
