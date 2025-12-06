const XLSX = require('xlsx');
const path = require('path');

const OUTPUT_FILE = path.join(__dirname, 'filtered_SUBTLEX_List.xlsx');

console.log('检查输出文件...\n');

const workbook = XLSX.readFile(OUTPUT_FILE);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

console.log('表头:', data[0]);
console.log(`\n总行数: ${data.length}`);
console.log(`数据行数: ${data.length - 1}`);

console.log('\n前5行数据示例:');
for (let i = 1; i <= Math.min(5, data.length - 1); i++) {
    console.log(`\n第${i}行:`);
    const row = data[i];
    data[0].forEach((header, index) => {
        console.log(`  ${header}: ${row[index]}`);
    });
}
