const path = require("path");
const webpack = require("webpack");

module.exports = env => {
    const plugins = [];
    if (env && env.substring(0, 4) === "prod") {
        plugins.push(
            new webpack.optimize.UglifyJsPlugin({
                compress: {
                    warnings: false,
                },
                output: {
                    comments: false,
                },
            })
        );
    }
    return {
        entry: {
            "main": "./" + path.relative(process.cwd(), path.join(__dirname, "src", "Calendar", "Extension.ts")),
            "calendarServices": "./" + path.relative(process.cwd(), path.join(__dirname, "src", "Calendar", "CalendarServices.ts")),
            "dialogs": "./" + path.relative(process.cwd(), path.join(__dirname, "src", "Calendar", "Dialogs.ts"))
        },
        output: {
            filename: path.relative(process.cwd(), path.join(__dirname, "dist", "[name].js")),
            libraryTarget: "amd",
        },
        externals: [
            {
                react: true,
                "react-dom": true,
            },
            /^TFS\//, // Ignore TFS/* since they are coming from VSTS host
            /^VSS\//, // Ignore VSS/* since they are coming from VSTS host
        ],
        resolve: {
            alias: { OfficeFabric: "../node_modules/office-ui-fabric-react/lib-amd" },
            extensions: [".ts", "tsx", ".js"],
        },
        module: {
            rules: [{ test: /\.tsx?$/, loader: "ts-loader" }, { test: /\.css$/, loader: "style-loader!css-loader" }],
        },
        plugins: plugins,
    };
};
