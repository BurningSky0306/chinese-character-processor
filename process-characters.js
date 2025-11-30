/**
 * 汉字属性处理脚本
 * 用于为汉字添加笔画、拼音、部首、结构、声母、韵母等属性
 */

const XLSX = require('xlsx');
const cnchar = require('cnchar');
const poly = require('cnchar-poly');
const radical = require('cnchar-radical');
const order = require('cnchar-order');

// 加载 cnchar 插件
cnchar.use(poly, radical, order);

console.log('==========================================');
console.log('   汉字属性处理脚本');
console.log('   用于心理学实验刺激材料准备');
console.log('==========================================\n');

// 读取 Excel 文件
const inputFileName = 'Chinese character list from 2.5 billion words corpus ordered by frequency.xlsx';
const outputFileName = 'Chinese character list with properties.xlsx';

console.log(`正在读取文件: ${inputFileName}...`);

let workbook;
try {
    workbook = XLSX.readFile(inputFileName);
} catch (error) {
    console.error(`错误: 无法读取文件 "${inputFileName}"`);
    console.error('请确保文件在当前目录下，并且文件名正确。');
    process.exit(1);
}

const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// 将工作表转换为 JSON 格式
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

console.log(`成功读取 ${data.length} 行数据\n`);
console.log('正在处理汉字属性...');

// 使用 cnchar.spellInfo 获取拼音详细信息(包括声母和韵母)

// 处理每一行数据
const headerRow = data[0];
const processedData = [];

// 添加新的列标题
const newHeaders = [
    ...headerRow,
    'stroke count',
    'pinyin',
    'radical',
    'structure',
    'Initial consonant',
    'final vowels'
];
processedData.push(newHeaders);

// 处理数据行(从第二行开始)
let processedCount = 0;
for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    if (!row || row.length === 0) {
        continue; // 跳过空行
    }
    
    // 获取汉字(第二列，索引为1)
    const character = row[1];
    
    if (!character || typeof character !== 'string') {
        // 如果没有汉字，添加空值
        processedData.push([...row, '', '', '', '', '', '']);
        continue;
    }
    
    try {
        // 1. 笔画数
        const strokeCount = cnchar.stroke(character);
        
        // 2. 拼音(多音字显示所有读音)
        const pinyinResult = cnchar.spell(character, 'tone', 'poly');
        const pinyin = Array.isArray(pinyinResult) ? pinyinResult.join(', ') : pinyinResult;
        
        // 3. 部首 - 使用 cnchar.radical() 并提取 radical 属性
        let radical = '';
        try {
            const radicalResult = cnchar.radical(character);
            if (Array.isArray(radicalResult) && radicalResult.length > 0) {
                radical = radicalResult[0].radical || '';
            }
        } catch (e) {
            radical = '';
        }
        
        // 4. 结构 - 使用 cnchar.radical() 的 struct 属性
        let structure = '';
        try {
            const radicalResult = cnchar.radical(character);
            if (Array.isArray(radicalResult) && radicalResult.length > 0) {
                structure = radicalResult[0].struct || '';
            }
        } catch (e) {
            structure = '';
        }
        
        // 5. 声母和韵母 - 使用 cnchar.spellInfo() 获取详细拼音信息
        let initial = '';
        let final = '';
        try {
            // 获取第一个读音(不带音调)
            const firstPinyin = Array.isArray(pinyinResult) ? pinyinResult[0] : pinyinResult;
            // 使用 spellInfo 获取声母和韵母
            const spellInfoResult = cnchar.spellInfo(firstPinyin);
            initial = spellInfoResult.initial || '';
            final = spellInfoResult.final || '';
        } catch (e) {
            initial = '';
            final = '';
        }
        
        // 添加处理后的行
        processedData.push([
            ...row,
            strokeCount,
            pinyin,
            radical,
            structure,
            initial,
            final
        ]);
        
        processedCount++;
        
        // 每处理1000个字符显示进度
        if (processedCount % 1000 === 0) {
            console.log(`已处理 ${processedCount} 个汉字...`);
        }
        
    } catch (error) {
        console.error(`处理字符 "${character}" 时出错:`, error.message);
        // 出错时添加空值
        processedData.push([...row, '', '', '', '', '', '']);
    }
}

console.log(`\n完成! 共处理 ${processedCount} 个汉字\n`);

// 创建新的工作表
const newWorksheet = XLSX.utils.aoa_to_sheet(processedData);

// 创建新的工作簿
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Characters');

// 写入文件
console.log(`正在保存到文件: ${outputFileName}...`);
XLSX.writeFile(newWorkbook, outputFileName);

console.log('\n==========================================');
console.log('   处理完成!');
console.log(`   输出文件: ${outputFileName}`);
console.log('==========================================\n');
console.log('新增的列包括:');
console.log('  - 第6列: stroke count (笔画数)');
console.log('  - 第7列: pinyin (拼音，多音字用逗号分隔)');
console.log('  - 第8列: radical (部首)');
console.log('  - 第9列: structure (结构)');
console.log('  - 第10列: Initial consonant (声母)');
console.log('  - 第11列: final vowels (韵母)');
console.log('\n您现在可以打开新生成的 Excel 文件查看结果。');
