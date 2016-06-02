var tsproject = require('tsproject');
var gulp = require('gulp');
var rjs = require('requirejs');

gulp.task('compile', function () {
    return tsproject.src('./').pipe(gulp.dest('./'));
});

gulp.task('bundle', [ 'compile' ], function () {
   return rjs.optimize({
        appDir: 'built/debug',
        baseUrl: './',
        dir: 'built/min',
        paths: {
            'VSS': 'empty:',
            'q': 'empty:',
            'jQuery': 'empty:',
            'TFS': 'empty:'
        },
        modules: [
            {
                name: 'Calendar/Extension'
            },
            {
                name: 'Calendar/Dialogs'
            },
            {
                name: 'Calendar/EventSources/FreeFormEventsSource'
            },
            {
                name: 'Calendar/EventSources/VSOCapacityEventSource'
            },
            {
                name: 'Calendar/EventSources/VSOIterationEventSource'
            },
        ],
    });
});

gulp.task('build', ['bundle']);