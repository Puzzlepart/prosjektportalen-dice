'use strict';
var gulp = require("gulp"),
    config = require('./@configuration.js'),
    runSequence = require("run-sequence"),
    powershell = require("./utils/powershell.js");

gulp.task("copyPnpTemplates", () => {
    return gulp.src(config.paths.templatesGlob)
        .pipe(gulp.dest(config.paths.templates_temp));
});

gulp.task("buildPnpTemplateFiles", ["copyPnpTemplates"], (done) => {
    powershell.execute("Build-PnP-Templates.ps1", "", done);
});