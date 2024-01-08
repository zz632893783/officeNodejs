// 南陵项目读取 excel 单元格内容，并批量生成 word 文档
const ExcelJS = require('exceljs');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { formatDate } = require('./util.js');
// 输出目录
const outputPath = './output/day/';
// 如果目录不存在，则创建目录
!fs.existsSync('./output/') && fs.mkdirSync('./output/');
!fs.existsSync(outputPath) && fs.mkdirSync(outputPath);
// 创建一个工作簿
const workbook = new ExcelJS.Workbook();
// 每日计划问题缓存列表
const plans = [];
// 加载工作表
workbook.xlsx.readFile('./input/dayData.xlsx').then(() => {
    // 选择第一个工作表
    const worksheet = workbook.getWorksheet(1);
    worksheet.eachRow({ includeEmpty: false }, row => {
        // 今日计划
        let today = [];
        // 明日计划
        let tomorrow = [];
        // 可能存在的问题
        let question = [];
        row.eachCell({ includeEmpty: true }, (cell, col) => {
            // 读取单元格内容，每行的第 5,6,7 列为 “今日计划，明日计划，可能存在的问题” 列
            col == 5 && cell.value && (today = cell.value.richText.filter(n => n.text.length > 2).map(n => n.text));
            col == 6 && cell.value && (tomorrow = cell.value.richText.filter(n => n.text.length > 2).map(n => n.text));
            col == 7 && cell.value && (question = cell.value.richText.filter(n => n.text.length > 2).map(n => n.text));
        });
        plans.push({ today, tomorrow, question });
    });
    const workDays = new Array(100).fill().map((n, i) => new Date(2000, 0, 1 + i));
    // workDays.splice(1, 1, new Date());
    plans.forEach((plan, i) => {
        const date = workDays[i];
        // 读取 word 模板文件
        const content = fs.readFileSync('./input/dayTemplate.docx', 'binary');
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
        fs.writeFileSync(`${ outputPath }南陵县新型智慧城市建设工程项目-城运中心子项目- ${ formatDate(date, 'yyyy-MM-dd') }.docx`, buf);
    });
});
