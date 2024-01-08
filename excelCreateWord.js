// 南陵项目读取 excel 单元格内容，并批量生成 word 文档
const ExcelJS = require('exceljs');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { formatDate } = require('./util.js');

// 创建一个工作簿
const workbook = new ExcelJS.Workbook();
// 每日计划问题缓存列表
const plans = [];
// 加载工作表
workbook.xlsx.readFile('./dataSource.xlsx').then(() => {
    // 选择第一个工作表
    const worksheet = workbook.getWorksheet(1);
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        // 今日计划
        let today = [];
        // 明日计划
        let tomorrow = [];
        // 可能存在的问题
        let question = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            // 输出单元格内容
            colNumber == 5 && cell.value && (today = cell.value.richText.filter(n => n.text.length > 2).map(n => n.text));
            colNumber == 6 && cell.value && (tomorrow = cell.value.richText.filter(n => n.text.length > 2).map(n => n.text));
            colNumber == 7 && cell.value && (question = cell.value.richText.filter(n => n.text.length > 2).map(n => n.text));
        });
        plans.push({ today, tomorrow, question });
    });
    const date = new Date(2023, 0, 1);
    plans.forEach(plan => {
        // 读取 word 模板文件
        const content = fs.readFileSync('./origin.docx', 'binary');
        // 创建 word 编辑模块
        const doc = new Docxtemplater(new PizZip(content), { paragraphLoop: true, linebreaks: true });
        // word 模板替换
        doc.render({
            date: formatDate(date, 'yyyy-MM-dd'),
            today: plan.today.length ? plan.today.join('\n') : '无',
            tomorrow: plan.tomorrow.length ? plan.tomorrow.join('\n') : '无',
            question: plan.question.length ? plan.question.join('\n') : '无'
        });
        const buf = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
        fs.writeFileSync(`./output/南陵县新型智慧城市建设工程项目-城运中心子项目- ${ formatDate(date, 'yyyy-MM-dd') }.docx`, buf);
        date.setDate(date.getDate() + 1);
    });
});
