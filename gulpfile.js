var tsproject = require('tsproject');
var gulp = require('gulp');
var rjs = require('requirejs');

gulp.task('compile', function () {
    tsproject.src('./')
        .pipe(gulp.dest('./'));
});

gulp.task('bundle', function () {
   rjs.optimize({
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

gulp.task('build', [
    'compile',
    'bundle'
]);