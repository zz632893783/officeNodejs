// 南陵项目读取 excel 单元格内容，并批量生成 word 文档
const ExcelJS = require('exceljs');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { formatDate } = require('./util.js');
// 输出目录
const outputPath = './output/week/';
// 如果目录不存在，则创建目录
!fs.existsSync('./output/') && fs.mkdirSync('./output/');
!fs.existsSync(outputPath) && fs.mkdirSync(outputPath);
// 创建一个工作簿
const workbook = new ExcelJS.Workbook();
// 每日计划问题缓存列表
const plans = [];
// 加载工作表
workbook.xlsx.readFile('./input/weekData.xlsx').then(() => {
    // 选择第一个工作表
    const worksheet = workbook.getWorksheet(1);
    worksheet.eachRow({ includeEmpty: false }, row => {
        // 本周计划
        let thisWeek = [];
        // 下周计划
        let nextWeek = [];
        // 可能存在的问题
        let question = [];
        row.eachCell({ includeEmpty: true }, (cell, col) => {
            col == 1 && cell.value && (thisWeek = cell.value.richText.filter(n => n.text.length > 2).map(n => n.text));
            col == 2 && cell.value && (nextWeek = cell.value.richText.filter(n => n.text.length > 2).map(n => n.text));
            col == 3 && cell.value && (question = cell.value.richText.filter(n => n.text.length > 2).map(n => n.text));
        });
        plans.push({ thisWeek, nextWeek, question });
    });
    const date = new Date(2023, 4, 1);
    plans.forEach(plan => {
        // 读取 word 模板文件
        const content = fs.readFileSync('./input/weekTemplate.docx', 'binary');
        // 创建 word 编辑模块
        const doc = new Docxtemplater(new PizZip(content), { paragraphLoop: true, linebreaks: true });
        // word 模板替换
        const startDate = new Date(date);
        const endDate = new Date(date);
        endDate.setDate(endDate.getDate() + 6);
        const start = formatDate(startDate, 'yyyy年MM月dd日');
        const end = formatDate(endDate, 'yyyy年MM月dd日');
        doc.render({
            start,
            end,
            thisWeek: plan.thisWeek.length ? plan.thisWeek.join('\n') : '无',
            nextWeek: plan.nextWeek.length ? plan.nextWeek.join('\n') : '无',
            question: plan.question.length ? plan.question.join('\n') : '无'
        });
        const buf = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
        fs.writeFileSync(`${ outputPath }南陵县新型智慧城市建设工程项目-城运中心子项目-施工周报（${ formatDate(startDate, 'yyyyMMdd') }-${ formatDate(endDate, 'yyyyMMdd') }）.docx`, buf);
        date.setDate(date.getDate() + 7);
    });
});
