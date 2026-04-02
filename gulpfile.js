'use strict';

const build = require('@microsoft/sp-build-web');

build.sass.setConfig({ tryToUseNpmModule: true });

// Disable sass task
build.sass.enabled = false;

build.initialize(require('gulp'));
