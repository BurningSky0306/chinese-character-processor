const XLSX = require('xlsx');
const path = require('path');

// 文件路径
const CHARACTER_FILE = path.join(__dirname, 'CHARACTER_List.xlsx');
const SUBTLEX_FILE = path.join(__dirname, 'SUBTLEX_List.xlsx');

console.log('开始处理Excel文件...');

// 读取两个Excel文件
const characterWorkbook = XLSX.readFile(CHARACTER_FILE);
const subtlexWorkbook = XLSX.readFile(SUBTLEX_FILE);

// 获取第一个工作表
const characterSheet = characterWorkbook.Sheets[characterWorkbook.SheetNames[0]];
const subtlexSheet = subtlexWorkbook.Sheets[subtlexWorkbook.SheetNames[0]];

// 将工作表转换为JSON格式（保留表头）
const characterData = XLSX.utils.sheet_to_json(characterSheet, { header: 1 });
const subtlexData = XLSX.utils.sheet_to_json(subtlexSheet, { header: 1 });

console.log(`CHARACTER_List.xlsx 总行数: ${characterData.length}`);
console.log(`SUBTLEX_List.xlsx 总行数: ${subtlexData.length}`);

// 创建CHARACTER_List中字符的Set（用于快速查找）
// 从B列（索引1）提取所有字符
const characterSet = new Set();
for (let i = 1; i < characterData.length; i++) {
    const character = characterData[i][1]; // B列（索引1）
    if (character) {
        characterSet.add(character.toString().trim());
    }
}

console.log(`CHARACTER_List中的字符数: ${characterSet.size}`);

// 在SUBTLEX的表头添加"existence"列
subtlexData[0][7] = 'existence'; // H列（索引7）

let existCount = 0;
let notExistCount = 0;

// 遍历SUBTLEX_List中的每个字符（从第2行开始）
for (let i = 1; i < subtlexData.length; i++) {
    const word = subtlexData[i][0]; // A列（索引0）
    
    if (!word) {
        subtlexData[i][7] = 0;
        continue;
    }
    
    const wordStr = word.toString().trim();
    
    // 检查该字符是否存在于CHARACTER_List中
    if (characterSet.has(wordStr)) {
        subtlexData[i][7] = 1;
        existCount++;
    } else {
        subtlexData[i][7] = 0;
        notExistCount++;
    }
    
    if ((i % 1000) === 0) {
        console.log(`已处理 ${i} 行...`);
    }
}

console.log(`\n处理完成！`);
console.log(`存在于CHARACTER_List中的字符: ${existCount} 个 (标记为1)`);
console.log(`不存在于CHARACTER_List中的字符: ${notExistCount} 个 (标记为0)`);

// 将修改后的数据转换回工作表
const newSheet = XLSX.utils.aoa_to_sheet(subtlexData);

// 更新工作簿
subtlexWorkbook.Sheets[subtlexWorkbook.SheetNames[0]] = newSheet;

// 保存修改后的文件
XLSX.writeFile(subtlexWorkbook, SUBTLEX_FILE);

console.log(`\nSUBTLEX_List.xlsx 已更新保存！`);
console.log(`H列 "existence" 已添加，值为 1（存在）或 0（不存在）`);
