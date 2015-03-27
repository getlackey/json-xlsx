/*jslint node:true */
'use strict';


module.exports = function browserify(grunt) {
    // Load task
    grunt.loadNpmTasks('grunt-browserify');

    // Options
    return {
        build: {
            files: grunt.file.expandMapping('lib/index.js', 'build/', {
                flatten: true,
                ext: '.js'
            }),
            options: {
                browserifyOptions: {
                    standalone: 'JsonXlsx'
                },
                watch: false,
                keepAlive: false,
                debug: false,
                require: []
            }
        }
    };
};