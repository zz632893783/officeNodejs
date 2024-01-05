// 创建 word 文档，并保存在本地
const docx = require('docx');
const fs = require('fs');
const { formatDate } = require('./util.js');

const { Document, Packer, Paragraph, TextRun } = docx;
const doc = new Document({
    sections: [
        {
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun('Hello, '),
                        new TextRun({ text: 'world!', bold: true })
                    ]
                })
            ]
        }
    ]
});

Packer.toBuffer(doc).then(buffer => {
    const fileName = `${ formatDate(new Date()) }.docx`;
    fs.writeFileSync(fileName, buffer);
});
