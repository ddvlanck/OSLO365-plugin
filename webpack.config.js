/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const VueLoaderPlugin = require("vue-loader/lib/plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { cacert: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const buildType = dev ? "dev" : "prod";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      search: "./src/taskpanes/search/search.ts",
      autoCheck: "./src/taskpanes/auto-check/auto-check.ts",
    },
    output: {
      devtoolModuleFilenameTemplate: "webpack:///[resource-path]?[loaders]",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js", ".vue", ".scss"],
      alias: {
        vue$: "vue/dist/vue.js",
      },
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-typescript"],
            },
          },
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: [
            {
              loader: "ts-loader",
              options: {
                appendTsSuffixTo: [/\.vue$/],
                transpileOnly: true,
              },
            },
          ],
        },
        {
          test: /\.vue$/,
          loader: "vue-loader",
          options: { esModule: true },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
        {
          test: /\.scss$/,
          use: ["vue-style-loader", "css-loader", "sass-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "search.html",
        template: "./src/taskpanes/search/search.html",
        chunks: ["polyfill", "search"],
      }),
      new HtmlWebpackPlugin({
        filename: "auto-check.html",
        template: "./src/taskpanes/auto-check/auto-check.html",
        chunks: ["polyfill", "autoCheck"],
      }),
      new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/taskpanes/functions.html",
        chunks: ["polyfill"],
      }),
      new HtmlWebpackPlugin({
        filename: "settings.html",
        template: "./src/taskpanes/settings/settings.html",
        chunks: ["polyfill"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "assets/icons/16x16/*",
            to: "assets/icons/16x16/[name][ext][query]",
          },
          {
            from: "assets/icons/32x32/*",
            to: "assets/icons/32x32/[name][ext][query]",
          },
          {
            from: "assets/icons/80x80/*",
            to: "assets/icons/80x80/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]." + buildType + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
          {
            from: "./src/data/oslo_terminology.json",
            to: "data/oslo_terminology.json",
          },
        ],
      }),
      new VueLoaderPlugin(),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      https: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
