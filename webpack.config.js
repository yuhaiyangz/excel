/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path  = require('path')
const urlDev = "https://localhost:3000/";
const urlProd = "https://excel.ccxindices.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

/* global require, module, process, __dirname */

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { cacert: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      functions: "./src/functions/functions.ts",
      taskpane: "./src/taskpane/taskpane.ts",
      commands: "./src/commands/commands.ts",
    },
    output: {
      devtoolModuleFilenameTemplate: "webpack:///[resource-path]?[loaders]",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    target: ['web', 'es5'],
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
          use: "ts-loader",
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          loader: "html-loader"
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        }
      ],
    },
    plugins: [
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: "./src/functions/functions.ts",
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "functions", "commands"],
        favicon:'assets/favicon.ico'
      }),
      new HtmlWebpackPlugin({
        filename: "login.html",
        template: "./src/login/index.html",
        chunks: ["login"],
        favicon:'assets/favicon.ico'
      }),
      new HtmlWebpackPlugin({
        filename: "functionSearch.html",
        template: "./src/functionSearch/index.html",
        chunks: ["functionSearch"],
        favicon:'assets/favicon.ico'
      }),
      new HtmlWebpackPlugin({
        filename: "help.html",
        template: "./src/help/index.html",
        chunks: [],
        favicon:'assets/favicon.ico'
      }),
      new HtmlWebpackPlugin({
        filename: "helpfile.html",
        template: "./src/help/helpfile.html",
        chunks: [],
        favicon:'assets/favicon.ico'
      }),
      new HtmlWebpackPlugin({
        filename: "join.html",
        template: "./src/help/join.html",
        chunks: [],
        favicon:'assets/favicon.ico'
      }),
      new HtmlWebpackPlugin({
        filename: "privacyPolicy.html",
        template: "./src/privacyPolicy/index.html",
        chunks: [],
        favicon:'assets/favicon.ico'
      }),
      new HtmlWebpackPlugin({
        filename: "getdata.html",
        template: "./src/getdata/index.html",
        chunks: [],
        favicon:'assets/favicon.ico'
      }),
      new CopyWebpackPlugin({
        patterns: [
          {from: path.join(__dirname, 'assets'),to: path.join(__dirname, 'dist/assets')},
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ]
      }),
    ],
    devServer: {
      static: [__dirname],
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      https: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
