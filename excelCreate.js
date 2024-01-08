// exceljs 创建 .xlsx 文件，并写入数据
const ExcelJS = require('exceljs');
const { formatDate } = require('./util.js');

// 创建一个工作簿
const workbook = new ExcelJS.Workbook();
// 创建一个工作表
const worksheet = workbook.addWorksheet('Sheet 1');
const rows = [
	['单元格1', '单元格2', '单元格3'],
	['单元格A', '单元格B', '单元格C'],
	['单元格甲', '单元格乙', '单元格丙']
];

rows.forEach(row => {
	const rowCells = worksheet.addRow(row);
	row.forEach((text, index) => {
		rowCells.getCell(index + 1).font = {
			name: 'Arial',
			family: Math.random() < 0.5 ? 1 : 2,
			italic: Math.random() < 0.5,
  			underline: Math.random() < 0.5,
			bold: Math.random() < 0.5,
			size: 10 + Math.round(Math.random() * 10),
			color: {
				argb: Math.round(Math.random() * 0xffffff).toString(16).padStart(6, 0).toUpperCase()
			}
		};
	});
});

// const column = table.getColumn(1);
// 保存文件到本地
const fileName = `${ formatDate(new Date()) }.xlsx`;
workbook.xlsx.writeFile(fileName)
    .then(() => console.log('Excel文件已保存成功！'))
    .catch(err => console.error('保存Excel文件时出现错误：', err));
