// 南陵项目读取 excel 单元格内容，并批量生成 word 文档
const ExcelJS = require('exceljs');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { formatDate } = require('./util.js');
// 输出目录
const outputPath = './output/month/';
// 如果目录不存在，则创建目录
!fs.existsSync('./output/') && fs.mkdirSync('./output/');
!fs.existsSync(outputPath) && fs.mkdirSync(outputPath);
// 创建一个工作簿
const workbook = new ExcelJS.Workbook();
// 每日计划问题缓存列表
const plans = [];
// 加载工作表
workbook.xlsx.readFile('./input/monthData.xlsx').then(() => {
    // 选择第一个工作表
    const worksheet = workbook.getWorksheet(1);
    worksheet.eachRow({ includeEmpty: false }, row => {
        // 本月计划
        let thisMonth = [];
        // 下月计划
        let nextMonth = [];
        row.eachCell({ includeEmpty: true }, (cell, col) => {
            col == 1 && cell.value && (thisMonth = cell.value.richText.filter(n => n.text.length > 2).map(n => n.text));
            col == 2 && cell.value && (nextMonth = cell.value.richText.filter(n => n.text.length > 2).map(n => n.text));
        });
        plans.push({ thisMonth, nextMonth });
    });
    const date = new Date(2023, 7);
    plans.forEach(plan => {
        // 读取 word 模板文件
        const content = fs.readFileSync('./input/monthTemplate.docx', 'binary');
        // 创建 word 编辑模块
        const doc = new Docxtemplater(new PizZip(content), { paragraphLoop: true, linebreaks: true });
        // word 模板替换
        doc.render({
            date: formatDate(date, 'yyyy年MM月'),
            thisMonth: plan.thisMonth.length ? plan.thisMonth.join('\n') : '无',
            nextMonth: plan.nextMonth.length ? plan.nextMonth.join('\n') : '无'
        });
        const buf = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
        fs.writeFileSync(`${ outputPath }南陵县新型智慧城市建设工程项目-城运中心子项目-施工月报（${ formatDate(date, 'MM月') }）.docx`, buf);
        date.setMonth(date.getMonth() + 1);
    });
});
