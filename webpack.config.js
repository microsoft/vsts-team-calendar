const path = require("path");
const fs = require("fs");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
    mode: "production", // "production" | "development" | "none"
    optimization: {
        minimize: true
    },
    entry: {
        Calendar: "./" + path.relative(process.cwd(), path.join(__dirname, "src", "Calendar.tsx"))
    },
    devtool: "inline-source-map",
    output: {
        filename: "[name].js"
    },
    resolve: {
        extensions: [".ts", ".tsx", ".js"],
        alias: {
            "azure-devops-extension-sdk": path.resolve("node_modules/azure-devops-extension-sdk")
        }
    },
    stats: {
        warnings: false
    },
    module: {
        rules: [
            {
                test: /\.tsx?$/,
                loader: "ts-loader"
            },
            {
                test: /\.scss$/,
                use: ["style-loader", "css-loader", "azure-devops-ui/buildScripts/css-variables-loader", "sass-loader"]
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
                test: /\.html$/,
                loader: "file-loader"
            }
        ]
    },
    plugins: [new CopyWebpackPlugin([{ from: "**/*.html", context: "src" }])]
};
