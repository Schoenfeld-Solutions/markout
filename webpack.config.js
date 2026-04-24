const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
const webpack = require("webpack");

module.exports = (_env, options) => {
  const env = _env ?? {};
  const isProduction = options.mode === "production";
  const includeTaskpaneMock =
    env.taskpaneMock === true || env.taskpaneMock === "true";

  const entry = {
    commands: "./src/commands/commands.ts",
    launchevent: "./src/launchevent/launchevent.ts",
    taskpane: "./src/taskpane/taskpane.tsx",
  };

  if (includeTaskpaneMock) {
    entry["taskpane-mock"] = "./src/taskpane/mock/taskpane-mock.tsx";
  }

  const plugins = [
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
  ];

  if (includeTaskpaneMock) {
    plugins.push(
      new HtmlWebpackPlugin({
        chunks: ["taskpane-mock"],
        filename: "taskpane-mock.html",
        template: "./src/taskpane/mock/taskpane-mock.html",
      })
    );
  }

  return {
    devtool: "source-map",
    entry,
    mode: isProduction ? "production" : "development",
    module: {
      rules: [
        {
          test: /\.tsx?$/,
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
      splitChunks: {
        chunks: "async",
        cacheGroups: {
          reactVendor: {
            test: /[\\/]node_modules[\\/](react|react-dom|scheduler)[\\/]/,
            name: "react-vendor",
            priority: 40,
            enforce: true,
            reuseExistingChunk: true,
          },
          fluentVendor: {
            test: /[\\/]node_modules[\\/](?:@fluentui|@griffel|@emotion|keyborg)[\\/]/,
            name: "fluent-vendor",
            priority: 30,
            enforce: true,
            reuseExistingChunk: true,
          },
        },
      },
    },
    output: {
      chunkFilename: "[name].chunk.js",
      clean: true,
      filename: "[name].js",
      path: path.resolve(__dirname, "dist"),
      publicPath: "",
    },
    plugins,
    resolve: {
      extensions: [".tsx", ".ts", ".html", ".js"],
    },
    target: "web",
  };
};
