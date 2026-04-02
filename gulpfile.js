'use strict';
const build = require('@microsoft/sp-build-web');
build.addSuppression(`Warning - [sass] The reference path is relative to the build file...`);
build.addSuppression(`Warning - [configure-webpack]`);

const gulp = require('gulp');

build.initialize(gulp);
