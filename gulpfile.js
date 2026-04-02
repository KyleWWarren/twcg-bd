'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The reference path is relative to the ` +
  `build file '${__dirname}' and not the build directory.`);

const getTasks = build.serial(
  build.preCopy,
  build.sass,
  build.tsc,
  build.postCopy,
  build.manifest,
  build.copyStaticAssets
);

var getTasks2 = build.parallel(
  build.clean
);

build.initialize(require('gulp'));