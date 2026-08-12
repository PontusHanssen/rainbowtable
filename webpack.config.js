/* eslint-disable no-undef */

const fs = require("fs");
const os = require("os");
const path = require("path");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

const urlDev = "https://localhost:3000/";
const urlProd = "https://pontushanssen.github.io/rainbowtable/"; // GitHub Pages, see .github/workflows/deploy.yml

// office-addin-dev-certs installs its CA into a trust path that does not exist on
// Fedora, and getHttpsServerOptions() retries that install on every start — which hangs
// the dev server. Once the certificates exist, read them straight off disk instead.
// Generate them with `npx office-addin-dev-certs install` (the CA install step may fail;
// the certificates are still written) and trust ca.crt in whichever browser you use.
async function getHttpsOptions() {
  const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
  const [ca, key, cert] = ["ca.crt", "localhost.key", "localhost.crt"].map((file) =>
    path.join(certDir, file)
  );

  if ([ca, key, cert].every((file) => fs.existsSync(file))) {
    return { ca: fs.readFileSync(ca), key: fs.readFileSync(key), cert: fs.readFileSync(cert) };
  }

  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|ttf|woff|woff2|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                // Match the origin with or without its trailing slash: AppDomain entries
                // carry no trailing slash, so a slash-terminated pattern leaves them
                // pointing at localhost in the built manifest.
                const from = urlDev.replace(/\/+$/, "");
                const to = urlProd.replace(/\/+$/, "");
                return content.toString().replace(new RegExp(from, "g"), to);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
