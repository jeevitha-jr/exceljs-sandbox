const Excel = require('exceljs');

const workbook = new Excel.Workbook();

workbook.creator = 'Lars Thorup';
workbook.lastModifiedBy = 'Lars Thorup';
workbook.created = new Date();
workbook.modified = new Date();

const sheet = workbook.addWorksheet('account');
sheet.columns = [
  {key: 'idName', header: 'id', style: {font: {bold: true, name: 'Calibri'}}},
  {key: 'displayName', header: 'name'}
];
const headerRow = sheet.getRow(1);
headerRow.eachCell(cell => {
  cell.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};
});
sheet.addRow({idName: 'lars', displayName: 'Lars Thorup'});
sheet.addRow({idName: 'finn', displayName: 'Finn Christensen'});

async function main () {
  await workbook.xlsx.writeFile('account.xlsx');
  console.log('Done');
}

main();