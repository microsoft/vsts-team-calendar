var tsproject = require('tsproject');
var gulp = require('gulp');
var shell = require('gulp-shell')

gulp.task( 'compile', function() {
    tsproject.src('./')
        .pipe(gulp.dest('./'));
});

gulp.task('bundle', shell.task([
    'r.js.cmd -o build.js'
]));

gulp.task('build', [
    'compile',
    'bundle'
]);