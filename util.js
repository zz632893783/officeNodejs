const fs = require('fs');
const path = require('path');
// 日期格式化
const formatDate = (date, format = 'yyyy-MM-dd_HH：mm：ss') => {
    const map = {
        'yyyy': date.getFullYear(),
        'MM': String(date.getMonth() + 1).padStart(2, '0'),
        'dd': String(date.getDate()).padStart(2, '0'),
        'HH': String(date.getHours()).padStart(2, '0'),
        'mm': String(date.getMinutes()).padStart(2, '0'),
        'ss': String(date.getSeconds()).padStart(2, '0'),
    };
    return format.replace(/yyyy|MM|dd|HH|mm|ss/g, match => map[match]);
};
// 读取某个目录下的所有文件
// rootPath 为根目录
// regExp 为文件名匹配的正则
// 用法 readDirAllFilePath('./root', /\.txt$/).then(pathList => console.log(pathList))
const readDirAllFilePath = async (rootPath, regExp = /./) => {
    const filePathCache = [];
    (function fn(dir) {
        const files = fs.readdirSync(dir);
        files.forEach(file => {
            const filePath = path.join(dir, file);
            const stat = fs.statSync(filePath);
            stat.isDirectory() ? fn(filePath) : (regExp.test(filePath) && filePathCache.push(filePath));
        });
    })(rootPath);
    return filePathCache;
}

module.exports = {
    formatDate,
    readDirAllFilePath
};