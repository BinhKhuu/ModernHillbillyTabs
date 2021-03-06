'use strict';
const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

const argv = build.rig.getYargs().argv;
const useCustomServe = argv['custom-serve'];
const fs = require("fs");
const workbenchApi = require("@microsoft/sp-webpart-workbench/lib/api");
const del = require('del');

if (useCustomServe) {
  build.tslintCmd.enabled = false;
  const ensureWorkbenchSubtask = build.subTask('ensure-workbench-task', function (gulp, buildOptions, done) {
    this.log('Creating workbench.html file...');
    try {
      workbenchApi.default["/workbench"]();
    } catch (e) { }

    done();
  });

  build.rig.addPostBuildTask(build.task('ensure-workbench', ensureWorkbenchSubtask));

  build.configureWebpack.mergeConfig({
    additionalConfiguration: (generatedConfiguration) => {
      fs.writeFileSync("./temp/_webpack_config.json", JSON.stringify(generatedConfiguration, null, 2));
      return generatedConfiguration;
    }
  });

} else {
  const deleteDefinitions = build.subTask('delete-scss-definitions-task', function (gulp, buildOptions, done) {
    del.sync(['src/**/*.scss.d.ts']);

    done();
  });

  build.rig.addPreBuildTask(build.task('delete-scss-definitions', deleteDefinitions));
}

build.initialize(require('gulp'));

