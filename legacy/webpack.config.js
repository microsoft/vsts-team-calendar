const path = require("path");
const webpack = require("webpack");
const ExtractTextPlugin = require("extract-text-webpack-plugin");
const extractSass = new ExtractTextPlugin({
    filename: path.relative(process.cwd(), path.join(__dirname, "dist", "css", "style.css"))
});

module.exports = env => {
    const plugins = [extractSass];
    if (env && env.substring(0, 4) === "prod") {
        plugins.push(
            new webpack.optimize.UglifyJsPlugin({
                compress: {
                    warnings: false
                },
                output: {
                    comments: false
                }
            })
        );
    }
    return {
        entry: {
            main: "./" + path.relative(process.cwd(), path.join(__dirname, "src", "Calendar", "Extension.tsx")),
            calendarServices: "./" + path.relative(process.cwd(), path.join(__dirname, "src", "Calendar", "CalendarServices.ts")),
            dialogs: "./" + path.relative(process.cwd(), path.join(__dirname, "src", "Calendar", "Dialogs.ts"))
        },
        output: {
            filename: path.relative(process.cwd(), path.join(__dirname, "dist", "js", "[name].js")),
            libraryTarget: "amd"
        },
        externals: [
            {
                react: true,
                "react-dom": true
            },
            // Ignore TFS/*, VSS/*, Favorites/* since they are coming from VSTS host
            /^TFS\//,
            /^VSS\//,
            /^Favorites\//
        ],
        resolve: {
            alias: { OfficeFabric: "../node_modules/office-ui-fabric-react/lib-amd" },
            extensions: [".ts", ".tsx", ".js"]
        },
        module: {
            rules: [
                { test: /\.tsx?$/, loader: "ts-loader" },
                {
                    test: /\.scss$/,
                    use: extractSass.extract({
                        use: [
                            {
                                loader: "css-loader",
                                options: { importLoaders: 1 }
                            },
                            {
                                loader: "sass-loader"
                            },
                            {
                                loader: "postcss-loader"
                            }
                        ],
                        fallback: "style-loader"
                    })
                },
                {
                    test: /\.css$/,
                    use: ["style-loader", "css-loader"]
                }
            ]
        },
        plugins: plugins
    };
};
