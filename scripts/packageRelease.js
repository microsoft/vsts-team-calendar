var exec = require("child_process").exec;

// Package extension
var command = `tfx extension create --rev-version --overrides-file configs/release.json --manifest-globs vss-extension.json --no-prompt --json`;
exec(command, (error, stdout) => {
    if (error) {
        console.error(`Could not create package: '${error}'`);
        return;
    }
     let output = JSON.parse(stdout);
    console.log(`Package created ${output.path}`);
}
);