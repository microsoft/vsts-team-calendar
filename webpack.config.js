const path = require("path");
const webpack = require("webpack");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
    mode: "production",
  target: "web",
  optimization: {
    minimize: true
},
  entry: {
    Calendar: "./" + path.relative(process.cwd(), path.join(__dirname, "src", "Calendar.tsx"))
  },
  output: {
    filename: "[name].js",
    path:  path.resolve(__dirname, 'dist')

  },
  devtool: "inline-source-map",
  devServer: {
    static: "./",
    hot: true,
    port: 8888,
    server: "https"
  },
  resolve: {
    extensions: [".ts", ".tsx", ".js"],
    alias: {
      "azure-devops-extension-sdk": path.resolve("node_modules/azure-devops-extension-sdk")
    },
    modules: [path.resolve("."), "node_modules"]
  },
  module: {
    rules: [
           {
        test: /\.tsx?$/,
        use: "ts-loader"
      },
      {
        test: /\.scss$/,
        use: [
          "style-loader",
          "css-loader",
          "sass-loader"
        ]
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"]
      },
      {
        test: /\.woff$/,
        use: [
          {
            loader: "base64-inline-loader"
          }
        ]
      },
      {
        test: /\.(png|svg|jpg|gif|html)$/,
        use: "file-loader"
      }
    ]
  },
  plugins: [
    new CopyWebpackPlugin({
      patterns: [
        { from: "**/*.html", to: "./", context: "src" },
        { from: "**/*.png", to: "./static/v2-images", context: "static/v2-images" },
        { from: "./azure-devops-extension.json", to: "azure-devops-extension.json" },
        { from: "./overview.md", to: "./" },
      ]
    })
  ]
};