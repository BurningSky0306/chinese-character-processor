const XLSX = require('xlsx');
const path = require('path');

// 文件路径
const CHARACTER_FILE = path.join(__dirname, 'CHARACTER_List.xlsx');
const SUBTLEX_FILE = path.join(__dirname, 'SUBTLEX_List.xlsx');

// 读取两个Excel文件
const characterWorkbook = XLSX.readFile(CHARACTER_FILE);
const subtlexWorkbook = XLSX.readFile(SUBTLEX_FILE);

// 获取第一个工作表
const characterSheet = characterWorkbook.Sheets[characterWorkbook.SheetNames[0]];
const subtlexSheet = subtlexWorkbook.Sheets[subtlexWorkbook.SheetNames[0]];

// 将工作表转换为JSON格式
const characterData = XLSX.utils.sheet_to_json(characterSheet, { header: 1 });
const subtlexData = XLSX.utils.sheet_to_json(subtlexSheet, { header: 1 });

console.log('=== CHARACTER_List.xlsx ===');
console.log('表头:', characterData[0]);
console.log('前5行数据:');
for (let i = 1; i <= Math.min(5, characterData.length - 1); i++) {
    console.log(`第${i+1}行:`, characterData[i][0], '类型:', typeof characterData[i][0]);
}

console.log('\n=== SUBTLEX_List.xlsx ===');
console.log('表头:', subtlexData[0]);
console.log('前5行数据:');
for (let i = 1; i <= Math.min(5, subtlexData.length - 1); i++) {
    console.log(`第${i+1}行:`, subtlexData[i][0], '类型:', typeof subtlexData[i][0]);
}

// 测试匹配
console.log('\n=== 测试匹配 ===');
const testChar = characterData[1][0];
console.log(`测试字符: "${testChar}" (类型: ${typeof testChar})`);

// 在SUBTLEX中查找
let found = false;
for (let i = 1; i < subtlexData.length; i++) {
    if (subtlexData[i][0] === testChar) {
        console.log(`在SUBTLEX第${i+1}行找到匹配: "${subtlexData[i][0]}"`);
        found = true;
        break;
    }
}
if (!found) {
    console.log('未找到匹配');
    // 尝试字符串比较
    for (let i = 1; i < Math.min(10, subtlexData.length); i++) {
        if (subtlexData[i][0] && subtlexData[i][0].toString() === testChar.toString()) {
            console.log(`使用toString()在SUBTLEX第${i+1}行找到匹配`);
            found = true;
            break;
        }
    }
}
