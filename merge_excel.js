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

// 创建SUBTLEX的字符索引（用于快速查找）
// key是字符，value是该字符所在的行索引
const subtlexMap = new Map();
for (let i = 1; i < subtlexData.length; i++) {
    const word = subtlexData[i][0]; // 第一列是Word
    if (word) {
        subtlexMap.set(word.toString().trim(), i);
    }
}

console.log(`SUBTLEX索引中的字符数: ${subtlexMap.size}`);

// 处理CHARACTER_List.xlsx的表头
// 添加SUBTLEX的列名到表头（从M列开始）
const subtlexHeaders = subtlexData[0]; // SUBTLEX的表头
if (subtlexHeaders && subtlexHeaders.length > 1) {
    // 将B到G列的表头添加到CHARACTER的表头后面
    for (let col = 1; col <= 6; col++) { // B列(索引1)到G列(索引6)
        characterData[0].push(subtlexHeaders[col] || '');
    }
}

let matchCount = 0;
let notFoundCount = 0;

// 遍历CHARACTER_List中的每个字符（从第2行开始）
for (let i = 1; i < characterData.length; i++) {
    const character = characterData[i][1]; // 第二列是character（索引1）
    
    if (!character) {
        console.log(`第 ${i + 1} 行：字符为空，跳过`);
        // 未找到该字符，在M到R列填充空值
        while (characterData[i].length < 18) {
            characterData[i].push('');
        }
        continue;
    }
    
    const characterStr = character.toString().trim();
    
    // 在SUBTLEX中查找该字符
    if (subtlexMap.has(characterStr)) {
        const subtlexRowIndex = subtlexMap.get(characterStr);
        const subtlexRow = subtlexData[subtlexRowIndex];
        
        // 确保当前行有足够的列（A到R列，即18列）
        while (characterData[i].length < 18) {
            characterData[i].push('');
        }
        
        // 将SUBTLEX的B到G列（索引1-6）复制到CHARACTER的M到R列（索引12-17）
        for (let col = 1; col <= 6; col++) {
            characterData[i][11 + col] = subtlexRow[col] !== undefined ? subtlexRow[col] : '';
        }
        
        matchCount++;
        if (matchCount % 100 === 0) {
            console.log(`已处理 ${matchCount} 个匹配的字符...`);
        }
    } else {
        // 未找到该字符，在M到R列填充空值
        while (characterData[i].length < 18) {
            characterData[i].push('');
        }
        notFoundCount++;
    }
}

console.log(`\n处理完成！`);
console.log(`匹配成功: ${matchCount} 个字符`);
console.log(`未找到: ${notFoundCount} 个字符`);

// 将修改后的数据转换回工作表
const newSheet = XLSX.utils.aoa_to_sheet(characterData);

// 更新工作簿
characterWorkbook.Sheets[characterWorkbook.SheetNames[0]] = newSheet;

// 保存修改后的文件
XLSX.writeFile(characterWorkbook, CHARACTER_FILE);

console.log(`\nCHARACTER_List.xlsx 已更新保存！`);
