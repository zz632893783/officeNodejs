const { readDirAllFilePath } = require('./util.js');

readDirAllFilePath('../vite-project/src', /\.vue$/).then(pathList => console.log(JSON.stringify(pathList)))