'use strict';

const build = require('@microsoft/sp-build-web');
const gulp = require('gulp');

build.addSuppression(
  `Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`
);

// Parent Next.js eslint packages conflict with SPFx lint; disable for reliable local builds.
build.lintCmd.enabled = false;

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);
  result.set('serve', result.get('serve-deprecated'));
  return result;
};

// Ensure AMD loc string modules are present under lib/ for webpack externals.
const copyLoc = build.subTask('copy-loc', function (gulpInstance, buildOptions, done) {
  return gulpInstance
    .src('src/webparts/**/loc/*.js')
    .pipe(gulpInstance.dest('lib/webparts'))
    .on('end', done);
});
build.rig.addPostBuildTask(copyLoc);

build.initialize(require('gulp'));
