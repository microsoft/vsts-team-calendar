const path = require("path");
const webpack = require("webpack");
const ExtractTextPlugin = require("extract-text-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const extractSass = new ExtractTextPlugin({
    filename: path.relative(process.cwd(), path.join(__dirname, "dist", "css", "style.css"))
});

module.exports = env => {
    
    const plugins = [];
    // if (env && env.substring(0, 4) === "prod") {
    //     plugins.push(
    //         new webpack.optimize.UglifyJsPlugin({
    //             compress: {
    //                 warnings: false
    //             },
    //             output: {
    //                 comments: false
    //             }
    //         })
    //     );
    // }
    return {
        mode: "development",
        entry: {
            main: "./" + path.relative(process.cwd(), path.join(__dirname, "src", "Calendar", "Extension.tsx")),
            calendarServices: "./" + path.relative(process.cwd(), path.join(__dirname, "src", "Calendar", "CalendarServices.ts")),
            dialogs: "./" + path.relative(process.cwd(), path.join(__dirname, "src", "Calendar", "Dialogs.ts"))
        },
        output: {
            filename: path.relative(process.cwd(), path.resolve(__dirname, "js", "[name].js")),
            // filename: '[name].js',
            // path: path.resolve(__dirname, 'dist', 'js'),
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
                // {
                //     test: /\.scss$/,
                //     use: extractSass.extract({
                //         use: [
                //             {
                //                 loader: "css-loader",
                //                 options: { importLoaders: 1 }
                //             },
                //             {
                //                 loader: "sass-loader"
                //             },
                //             {
                //                 loader: "postcss-loader"
                //             }
                //         ],
                //         fallback: "style-loader"
                //     })
                // },
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
                }
            ]
        },
        plugins: [
    new CopyWebpackPlugin({
      patterns: [
        { from: "**/*.html", to: "./dist", context: "static" },
        { from: "node_modules/vss-web-extension-sdk/lib", to: "./sdk"},
        { from: "node_modules/moment/min", to: "./lib/moment"},
        { from: "node_modules/jquery/dist", to: "./lib/jquery"},
        { from: "node_modules/fullcalendar/dist", to: "./lib/fullcalendar"}
        
      ]
    })
  ]
    };
};
