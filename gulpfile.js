'use strict';
const build = require('@microsoft/sp-build-web');
build.sass.setConfig({ tryToUseNpmModule: true });
build.sass.enabled = false;
build.addSuppression(/Warning - \[sass\]/gi);
build.initialize(require('gulp'));