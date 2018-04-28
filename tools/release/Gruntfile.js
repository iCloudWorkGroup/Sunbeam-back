module.exports = function(grunt) {
    'use strict';
    grunt.initConfig({
        pkg: grunt.file.readJSON('package.json'),
        bump: {
            options: {
                updateConfigs: ['pkg'],
                commitFIles: ['package.json', '../../CHANGELOG.md'],
                commitMessage: 'spreadsheet: release v%VERSION%',
                push: false
            }
        },
        conventionalChangelog: {
            options: {
                changelogOpts: {
                    // conventional-changelog options go here
                    preset: 'angular',
                    outputUnreleased: true,
                    releaseCount: 0
                },
                context: {
                    // context goes here
                },
                gitRawCommitsOpts: {
                    // git-raw-commits options go here
                },
                parserOpts: {
                    // conventional-commits-parser options go here
                },
                writerOpts: {
                    // conventional-changelog-writer options go here
                }
            },
            release: {
                src: '../../CHANGELOG.md'
            }
        }
    });
    require('load-grunt-tasks')(grunt);
    grunt.loadNpmTasks('grunt-conventional-changelog');
    grunt.registerTask('release', 'build new version info', function(type) {
        grunt.task.run([
            'bump:' + (type || 'patch') + ':bump-only',
            'conventionalChangelog',
            'bump-commit'
        ]);
    });
};