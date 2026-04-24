const fs = require("fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const os = require("os");
const path = require("path");
const webpack = require("webpack");

function readHttpsServerOptions() {
  const defaultCertificateDirectory = path.join(
    os.homedir(),
    ".office-addin-dev-certs"
  );
  const certificatePath =
    process.env.MARKOUT_DEV_TLS_CERT_PATH ??
    path.join(defaultCertificateDirectory, "localhost.crt");
  const keyPath =
    process.env.MARKOUT_DEV_TLS_KEY_PATH ??
    path.join(defaultCertificateDirectory, "localhost.key");

  if (!fs.existsSync(certificatePath) || !fs.existsSync(keyPath)) {
    throw new Error(
      "Local HTTPS requires MARKOUT_DEV_TLS_CERT_PATH and MARKOUT_DEV_TLS_KEY_PATH, or existing ~/.office-addin-dev-certs/localhost.crt and localhost.key files."
    );
  }

  return {
    cert: fs.readFileSync(certificatePath),
    key: fs.readFileSync(keyPath),
  };
}

module.exports = async (_env, options) => {
  const env = _env ?? {};
  const isProduction = options.mode === "production";
  const includeTaskpaneMock =
    env.taskpaneMock === true || env.taskpaneMock === "true";
  const useHttpDevServer =
    env.taskpaneHttps === false || env.taskpaneHttps === "false";
  const devServerServer =
    !isProduction && !useHttpDevServer
      ? {
          type: "https",
          options: readHttpsServerOptions(),
        }
      : "http";

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
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      server: devServerServer,
    },
  };
};
