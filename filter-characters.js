/**
 * 汉字过滤脚本
 * 用于过滤笔画数为零和多音字的汉字
 */

const XLSX = require('xlsx');
const cnchar = require('cnchar');
const poly = require('cnchar-poly');
const radical = require('cnchar-radical');
const order = require('cnchar-order');

// 加载 cnchar 插件
cnchar.use(poly, radical, order);

console.log('==========================================');
console.log('   汉字过滤脚本');
console.log('   过滤笔画数为零和多音字');
console.log('==========================================\n');

// 读取 Excel 文件
const inputFileName = 'Chinese character list with properties.xlsx';
const outputFileName = 'Chinese character list filtered.xlsx';

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
console.log('正在过滤数据...');

// 存储过滤后的数据
const filteredData = [];

// 添加表头（第一行）
const headerRow = data[0];
filteredData.push(headerRow);

// 统计信息
let totalProcessed = 0;
let excludedZeroStroke = 0;
let excludedPolyphone = 0;
let kept = 0;

// 处理数据行(从第二行开始)
for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    if (!row || row.length === 0) {
        continue; // 跳过空行
    }
    
    totalProcessed++;
    
    // 获取汉字(第二列，索引为1)
    const character = row[1];
    
    // 获取笔画数(第六列，索引为5)
    const strokeCount = row[5];
    
    // 检查1: 笔画数是否为0
    if (strokeCount === 0 || strokeCount === '0') {
        excludedZeroStroke++;
        continue; // 排除这一行
    }
    
    // 检查2: 是否为多音字
    if (character && typeof character === 'string' && character.length === 1) {
        try {
            const isPolyWord = cnchar.isPolyWord(character);
            if (isPolyWord) {
                excludedPolyphone++;
                continue; // 排除这一行
            }
        } catch (error) {
            // 如果检查出错，保留这一行
            console.warn(`检查字符 "${character}" 时出错: ${error.message}`);
        }
    }
    
    // 通过所有检查，保留这一行
    filteredData.push(row);
    kept++;
    
    // 每处理1000个字符显示进度
    if (totalProcessed % 1000 === 0) {
        console.log(`已处理 ${totalProcessed} 行...`);
    }
}

console.log(`\n过滤完成!`);
console.log(`==========================================`);
console.log(`统计信息:`);
console.log(`  总处理行数: ${totalProcessed}`);
console.log(`  排除笔画数为0的行: ${excludedZeroStroke}`);
console.log(`  排除多音字的行: ${excludedPolyphone}`);
console.log(`  保留的行数: ${kept}`);
console.log(`  保留比例: ${(kept / totalProcessed * 100).toFixed(2)}%`);
console.log(`==========================================\n`);

// 创建新的工作表
const newWorksheet = XLSX.utils.aoa_to_sheet(filteredData);

// 创建新的工作簿
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Filtered Characters');

// 写入文件
console.log(`正在保存到文件: ${outputFileName}...`);
XLSX.writeFile(newWorkbook, outputFileName);

console.log('\n==========================================');
console.log('   处理完成!');
console.log(`   输出文件: ${outputFileName}`);
console.log('==========================================\n');
console.log('过滤规则:');
console.log('  1. 已排除笔画数为 0 的汉字');
console.log('  2. 已排除多音字');
console.log('\n您现在可以打开新生成的 Excel 文件查看结果。');
