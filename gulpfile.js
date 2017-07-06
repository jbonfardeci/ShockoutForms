var gulp = require('gulp'),
    bower = require('bower'),
    uglify = require("gulp-uglify"),
    concat = require('gulp-concat'),
    typescript = require('gulp-typescript'),
    sass = require('gulp-sass'),
    replace = require('gulp-replace'),
    util = require('gulp-util');

var tsFiles = "TypeScript/**/*.ts",
    ver = '1.0.9';

// Compile and combine all TypeScript files in ts/ into www/js/appBundle.js
gulp.task('ts', function(){
    return gulp.src(tsFiles)
        .pipe(typescript(
            {
                noImplicitAny: false,
                noEmitOnError: true,
                removeComments: true,
                sourceMap: false,
                out: "ShockoutForms-" + ver + ".js",
                target: "es5"
            }))
        .pipe(gulp.dest("JavaScript"));
});

gulp.task('ver', ['ts'], function(){
    return gulp.src('Javascript/ShockoutForms-' + ver + '.js')
        .pipe(replace(
                /this\.version = '.*';/g, 
                function(match){
                    var newString = "this.version = '" + ver + "';";
                    console.log('found: ' + match);
                    console.log('updating with : ' + newString);
                    return newString;
                }
            ))
        .pipe(gulp.dest('Javascript'));
});

// Minify and uglify all JavaScript files in www/js/ to into www/js/appBundle.min.js
gulp.task('min', ['ver'], function(){
    // js
    return gulp.src('JavaScript/ShockoutForms-' + ver + '.js')
        .pipe(concat('ShockoutForms-' + ver + '.min.js'))
        .pipe(uglify())
        .pipe(gulp.dest('JavaScript/'));
});

// Compile all SASS (.scss) files to www/css/
gulp.task('sass', function () {
  return gulp.src('scss/**/*.scss')
    .pipe(sass({
        outputStyle: 'compressed'
    }))
    .pipe(sass().on('error', sass.logError))   
    .pipe(gulp.dest('css/'));
});

gulp.task('watch', function(){
    gulp.watch(tsFiles, ['ts']);
});

gulp.task('default', ['min', 'sass'], function(){
   console.log('--------------the build of Shockout is complete ------------>');
});