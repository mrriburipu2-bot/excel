/* eslint-disable no-undef */
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

// Where the add-in is served from in development. For production, change
// urlProd (or set ADDIN_URL) to your HTTPS host; manifest URLs are rewritten during the build.
const urlDev = "https://localhost:3000/";
const urlProd = process.env.ADDIN_URL || "https://www.contoso.com/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  return {
    devtool: dev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
      functions: "./src/functions/functions.ts",
      commands: "./src/commands/commands.ts",
    },
    output: { clean: true },
    resolve: { extensions: [".ts", ".js"], extensionAlias: { ".js": [".ts", ".js"] } },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: { loader: "ts-loader", options: { transpileOnly: true, compilerOptions: { noEmit: false } } },
        },
        { test: /\.(png|jpg|jpeg|gif|ico|svg)$/, type: "asset/resource", generator: { filename: "assets/[name][ext][query]" } },
      ],
    },
    plugins: [
      new CustomFunctionsMetadataPlugin({ output: "functions.json", input: "./src/functions/functions.ts" }),
      // Shared runtime: one page hosts the task pane, the custom functions and the ribbon commands.
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane", "functions", "commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/*", to: "assets/[name][ext][query]" },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              return dev ? content : content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            },
          },
        ],
      }),
    ],
    devServer: {
      static: { directory: "dist", publicPath: "/" },
      headers: { "Access-Control-Allow-Origin": "*" },
      server: { type: "https", options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions() },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };
};
