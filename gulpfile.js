'use strict';
const build = require('@microsoft/sp-build-web');
build.sass.setConfig({ tryToUseNpmModule: true });
build.sass.enabled = false;
build.addSuppression(/Warning - \[sass\]/gi);
build.addSuppression(/Warning - \[lint\]/gi);
const buildLint = build.subTask('lint', function(gulp, buildConfig, done) { done(); });
build.task('lint', buildLint);
build.initialize(require('gulp'));