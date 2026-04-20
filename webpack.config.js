const devCerts = require("office-addin-dev-certs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
const webpack = require("webpack");

module.exports = async (_env, options) => {
  const isProduction = options.mode === "production";
  const httpsOptions =
    !isProduction && options.https === undefined
      ? await devCerts.getHttpsServerOptions()
      : options.https;

  return {
    devtool: "source-map",
    entry: {
      commands: "./src/commands/commands.ts",
      launchevent: "./src/launchevent/launchevent.ts",
      taskpane: "./src/taskpane/taskpane.ts",
    },
    mode: isProduction ? "production" : "development",
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "ts-loader",
          },
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/i,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext]",
          },
        },
      ],
    },
    optimization: {
      runtimeChunk: false,
    },
    output: {
      chunkFilename: "[name].chunk.js",
      clean: true,
      filename: "[name].js",
      path: path.resolve(__dirname, "dist"),
    },
    plugins: [
      new webpack.DefinePlugin({
        "process.env.NODE_DEBUG": JSON.stringify(""),
      }),
      new HtmlWebpackPlugin({
        chunks: ["commands", "launchevent"],
        filename: "commands.html",
        template: "./src/commands/commands.html",
      }),
      new HtmlWebpackPlugin({
        chunks: ["taskpane"],
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
      }),
    ],
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    target: "web",
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      server: {
        type: "https",
        options: httpsOptions,
      },
    },
  };
};
