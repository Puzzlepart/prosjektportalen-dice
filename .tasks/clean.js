'use strict';
var gulp = require("gulp"),
    clean = require('gulp-clean'),
    runSequence = require("run-sequence"),
    config = require('./@configuration.js');

gulp.task("clean", done => {
    return gulp.src([config.paths.dist, config.paths.templates_temp], { read: false }).pipe(clean());
});