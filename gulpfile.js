var gulp = require('gulp'),
    bower = require('bower'),
    uglify = require("gulp-uglify"),
    concat = require('gulp-concat'),
    typescript = require('gulp-typescript'),
    sass = require('gulp-sass');

var tsFiles = "TypeScript/**/*.ts",
    version = '1.0.9';

// Compile and combine all TypeScript files in ts/ into www/js/appBundle.js
gulp.task('ts', function(){
    return gulp.src(tsFiles)
        .pipe(typescript({
            noImplicitAny: false,
            noEmitOnError: true,
            removeComments: true,
            sourceMap: false,
            out: "ShockoutForms-" + version + ".js",
            target: "es5"
        })).pipe(gulp.dest("JavaScript"));
});

// Minify and uglify all JavaScript files in www/js/ to into www/js/appBundle.min.js
gulp.task('min', ['ts'], function(){
    // js
    return gulp.src('JavaScript/ShockoutForms-' + version + '.js')
        .pipe(concat('ShockoutForms-' + version + '.min.js'))
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