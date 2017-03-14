var gulp = require('gulp');
var bower = require('bower');
var uglify = require("gulp-uglify");
var concat = require('gulp-concat');
var typescript = require('gulp-typescript');
var sass = require('gulp-sass');
var tsFiles = "TypeScript/**/*.ts";

gulp.task('default', function(){});

// Compile and combine all TypeScript files in ts/ into www/js/appBundle.js
gulp.task('ts', function(){
    gulp.src(tsFiles)
        .pipe(typescript({
            noImplicitAny: false,
            noEmitOnError: true,
            removeComments: true,
            sourceMap: false,
            out: "ShockoutForms-1.0.8.js",
            target: "es5"
        })).pipe(gulp.dest("JavaScript"));
});

// Minify and uglify all JavaScript files in www/js/ to into www/js/appBundle.min.js
gulp.task('min', function(){
    // js
    gulp.src('JavaScript/ShockoutForms-1.0.8.js')
        .pipe(concat('ShockoutForms-1.0.8.min.js'))
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
    gulp.watch(tsFiles, ['TypeScript']);
});