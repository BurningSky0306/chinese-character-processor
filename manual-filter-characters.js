/**
 * 人工审核后汉字过滤脚本
 * 基于 Chinese character list filtered and manual delete.xlsx 进行处理
 */

const XLSX = require('xlsx');
const cnchar = require('cnchar');
const poly = require('cnchar-poly');
const radical = require('cnchar-radical');
const order = require('cnchar-order');

// 加载 cnchar 插件
cnchar.use(poly, radical, order);

console.log('==========================================');
console.log('   人工审核后汉字过滤脚本');
console.log('   提取声调并过滤数字同音字');
console.log('==========================================\n');

// 读取 Excel 文件
const inputFileName = 'Chinese character list filtered and manual delete.xlsx';
const outputFileName = 'Chinese character list final manual.xlsx';

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
console.log('正在处理数据...');

// 提取声调的函数（支持轻声）
function extractTone(pinyinWithTone) {
    if (!pinyinWithTone || typeof pinyinWithTone !== 'string') {
        return 0;
    }
    
    try {
        // 使用 cnchar.spellInfo 来获取声调信息
        const info = cnchar.spellInfo(pinyinWithTone);
        // cnchar 中 tone 为 0 表示轻声，我们转换为 5
        return info.tone === 0 ? 5 : info.tone;
    } catch (error) {
        // 如果解析失败，尝试判断是否有音调符号
        const hastoneMark = /[āáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜ]/i.test(pinyinWithTone);
        if (!hastoneMark) {
            // 没有音调符号，很可能是轻声
            return 5;
        }
        return 0;
    }
}

// 去除拼音音调的函数
function removeTone(pinyinWithTone) {
    if (!pinyinWithTone || typeof pinyinWithTone !== 'string') {
        return '';
    }
    
    try {
        const info = cnchar.spellInfo(pinyinWithTone);
        return info.spell.toLowerCase();
    } catch (error) {
        // 如果失败，手动去除音调符号
        return pinyinWithTone
            .toLowerCase()
            .replace(/[āáǎà]/g, 'a')
            .replace(/[ēéěè]/g, 'e')
            .replace(/[īíǐì]/g, 'i')
            .replace(/[ōóǒò]/g, 'o')
            .replace(/[ūúǔù]/g, 'u')
            .replace(/[ǖǘǚǜ]/g, 'v');
    }
}

// 数字拼音对照表 (0-10)
const numberPinyinMap = [
    { pinyin: 'ling', tone: 2 },  // 0 - 零
    { pinyin: 'yi', tone: 1 },    // 1 - 一
    { pinyin: 'er', tone: 4 },    // 2 - 二
    { pinyin: 'san', tone: 1 },   // 3 - 三
    { pinyin: 'si', tone: 4 },    // 4 - 四
    { pinyin: 'wu', tone: 3 },    // 5 - 五
    { pinyin: 'liu', tone: 4 },   // 6 - 六
    { pinyin: 'qi', tone: 1 },    // 7 - 七
    { pinyin: 'ba', tone: 1 },    // 8 - 八
    { pinyin: 'jiu', tone: 3 },   // 9 - 九
    { pinyin: 'shi', tone: 2 }    // 10 - 十
];

// 检查是否与数字拼音重复
function isNumberPinyin(pinyin, tone) {
    return numberPinyinMap.some(num => 
        num.pinyin === pinyin && num.tone === tone
    );
}

// 存储过滤后的数据
const filteredData = [];

// 添加表头（第一行），并添加新的 tone 列
const headerRow = data[0];
const newHeaders = [
    ...headerRow,
    'tone'  // 新增第12列：声调 (1,2,3,4,5)
];
filteredData.push(newHeaders);

// 统计信息
let totalProcessed = 0;
let excludedByNumberPinyin = 0;
let kept = 0;
let lightToneCount = 0;

// 处理数据行(从第二行开始)
for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    if (!row || row.length === 0) {
        continue; // 跳过空行
    }
    
    totalProcessed++;
    
    // 获取汉字(第二列，索引为1)
    const character = row[1];
    
    // 获取拼音(第七列，索引为6)
    const pinyinWithTone = row[6];
    
    // 提取声调和无声调拼音
    const tone = extractTone(pinyinWithTone);
    const pinyinWithoutTone = removeTone(pinyinWithTone);
    
    // 统计轻声字
    if (tone === 5) {
        lightToneCount++;
        console.log(`发现轻声字: ${character} (${pinyinWithTone}) - 声调: ${tone}`);
    }
    
    // 检查: 是否与数字拼音重复（注意：轻声字不会与数字重复，因为数字都有明确声调）
    if (tone !== 5 && isNumberPinyin(pinyinWithoutTone, tone)) {
        excludedByNumberPinyin++;
        console.log(`排除数字同音字: ${character} (${pinyinWithTone}) - 拼音: ${pinyinWithoutTone}, 声调: ${tone}`);
        continue;
    }
    
    // 通过所有检查，保留这一行并添加声调列
    const newRow = [...row, tone];
    filteredData.push(newRow);
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
console.log(`  发现轻声字: ${lightToneCount}`);
console.log(`  排除数字同音字的行: ${excludedByNumberPinyin}`);
console.log(`  保留的行数: ${kept}`);
console.log(`  保留比例: ${(kept / totalProcessed * 100).toFixed(2)}%`);
console.log(`==========================================\n`);

// 创建新的工作表
const newWorksheet = XLSX.utils.aoa_to_sheet(filteredData);

// 创建新的工作簿
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Final Manual Characters');

// 写入文件
console.log(`正在保存到文件: ${outputFileName}...`);
XLSX.writeFile(newWorkbook, outputFileName);

console.log('\n==========================================');
console.log('   处理完成!');
console.log(`   输出文件: ${outputFileName}`);
console.log('==========================================\n');
console.log('处理内容:');
console.log('  1. 新增第12列 "tone" (声调: 1/2/3/4/5)');
console.log('     - 1, 2, 3, 4 表示一声、二声、三声、四声');
console.log('     - 5 表示轻声');
console.log('  2. 排除与数字0-10拼音声调相同的汉字:');
console.log('     - 零(líng) 一(yī) 二(èr) 三(sān) 四(sì) 五(wǔ)');
console.log('     - 六(liù) 七(qī) 八(bā) 九(jiǔ) 十(shí)');
console.log('\n您现在可以打开新生成的 Excel 文件查看结果。');
