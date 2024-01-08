// 读取 word 文件，修改后保存在本地
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { formatDate, readDirAllFilePath } = require('./util.js');

// 读取的文件路径
const inputFilePath = './origin.docx';
// 保存的文件路径
const outputFilePath = `./${ formatDate(new Date()) }.docx`;

const content = fs.readFileSync(inputFilePath, 'binary');
const zip = new PizZip(content);
const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
// 源文档中，被 {} 包裹的文字，例如 { first_name } 会被替换为 render 传参中的值
// 如下
doc.render({
    date: '我是你爹',
    content: '内容内容内容内容内容内容内容'
});

const buf = doc.getZip().generate({
    type: 'nodebuffer',
    compression: 'DEFLATE',
});
fs.writeFileSync(outputFilePath, buf);
