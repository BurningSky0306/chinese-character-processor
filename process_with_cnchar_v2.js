const XLSX = require('xlsx');
const path = require('path');

// 导入cnchar及其插件
const cnchar = require('cnchar');
const poly = require('cnchar-poly');
const order = require('cnchar-order');
const trad = require('cnchar-trad');
const radical = require('cnchar-radical');
const info = require('cnchar-info');

// 使用插件
cnchar.use(poly, order, trad, radical, info);

// 声调符号转数字映射
const toneMap = {
    // 一声
    'ā': 'a1', 'ē': 'e1', 'ī': 'i1', 'ō': 'o1', 'ū': 'u1', 'ǖ': 'v1',
    'Ā': 'A1', 'Ē': 'E1', 'Ī': 'I1', 'Ō': 'O1', 'Ū': 'U1', 'Ǖ': 'V1',
    // 二声
    'á': 'a2', 'é': 'e2', 'í': 'i2', 'ó': 'o2', 'ú': 'u2', 'ǘ': 'v2',
    'Á': 'A2', 'É': 'E2', 'Í': 'I2', 'Ó': 'O2', 'Ú': 'U2', 'Ǘ': 'V2',
    // 三声
    'ǎ': 'a3', 'ě': 'e3', 'ǐ': 'i3', 'ǒ': 'o3', 'ǔ': 'u3', 'ǚ': 'v3',
    'Ǎ': 'A3', 'Ě': 'E3', 'Ǐ': 'I3', 'Ǒ': 'O3', 'Ǔ': 'U3', 'Ǚ': 'V3',
    // 四声
    'à': 'a4', 'è': 'e4', 'ì': 'i4', 'ò': 'o4', 'ù': 'u4', 'ǜ': 'v4',
    'À': 'A4', 'È': 'E4', 'Ì': 'I4', 'Ò': 'O4', 'Ù': 'U4', 'Ǜ': 'V4'
};

// 将带声调符号的拼音转换为数字形式
function convertToneToNumber(pinyin) {
    let result = pinyin;
    let tone = 0;
    
    for (const [symbol, replacement] of Object.entries(toneMap)) {
        if (pinyin.includes(symbol)) {
            result = pinyin.replace(symbol, replacement[0]);
            tone = parseInt(replacement[replacement.length - 1]);
            break;
        }
    }
    
    return { pinyin: result.toLowerCase(), tone };
}

// 提取声母和韵母的辅助函数
function extractInitialAndFinal(pinyin) {
    const pure = pinyin.toLowerCase();
    
    const initials = ['zh', 'ch', 'sh', 'b', 'p', 'm', 'f', 'd', 't', 'n', 'l', 
                     'g', 'k', 'h', 'j', 'q', 'x', 'r', 'z', 'c', 's', 'y', 'w'];
    
    let initial = '';
    let final = pure;
    
    for (const init of initials) {
        if (pure.startsWith(init)) {
            initial = init;
            final = pure.substring(init.length);
            break;
        }
    }
    
    if (!initial) {
        final = pure;
    }
    
    return { initial: initial || '', final: final || pure };
}

// 文件路径
const SUBTLEX_FILE = path.join(__dirname, 'SUBTLEX_List.xlsx');
const OUTPUT_FILE = path.join(__dirname, 'filtered_SUBTLEX_List.xlsx');

console.log('开始处理Excel文件...\n');

// 读取Excel文件
const workbook = XLSX.readFile(SUBTLEX_FILE);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

console.log('原始表头:', data[0]);
console.log(`原始总行数: ${data.length}\n`);

// 添加新列到表头
const newColumns = [
    'stroke count',
    'pinyin',
    'radical',
    'structure',
    'Initial consonant',
    'final vowels',
    'tone'
];

data[0] = [...data[0], ...newColumns];

// 数字0-10的拼音和声调对照
const numberPinyinTone = [
    { pinyin: 'ling', tone: 2 },
    { pinyin: 'yi', tone: 1 },
    { pinyin: 'er', tone: 4 },
    { pinyin: 'san', tone: 1 },
    { pinyin: 'si', tone: 4 },
    { pinyin: 'wu', tone: 3 },
    { pinyin: 'liu', tone: 4 },
    { pinyin: 'qi', tone: 1 },
    { pinyin: 'ba', tone: 1 },
    { pinyin: 'jiu', tone: 3 },
    { pinyin: 'shi', tone: 2 }
];

const rowsToDelete = new Set();
let polyPhoneCount = 0;
let neutralToneCount = 0;
let numberMatchCount = 0;
let noRadicalCount = 0;

console.log('开始处理每个字符...\n');

// 处理每一行
for (let i = 1; i < data.length; i++) {
    const word = data[i][0];
    
    if (!word || typeof word !== 'string' || word.length !== 1) {
        rowsToDelete.add(i);
        continue;
    }
    
    const char = word;
    
    try {
        // 获取笔画数
        const strokeCount = cnchar.stroke(char);
        
        // 获取拼音(带声调)，poly参数用于多音字
        const pinyinWithTone = cnchar.spell(char, 'tone', 'poly');
        
        // 获取部首（返回可能是对象或数组）
        const radicalResult = cnchar.radical(char);
        let radicalValue = '';
        if (typeof radicalResult === 'string') {
            radicalValue = radicalResult;
        } else if (Array.isArray(radicalResult)) {
            radicalValue = radicalResult[0] || '';
        } else if (radicalResult && typeof radicalResult === 'object') {
            radicalValue = radicalResult[0] || radicalResult.toString();
        }
        
        // 获取字形结构
        const structureResult = cnchar.info(char, 'struct');
        let structureValue = '';
        if (typeof structureResult === 'string') {
            structureValue = structureResult;
        } else if (Array.isArray(structureResult)) {
            structureValue = structureResult[0] || '';
        } else if (structureResult && typeof structureResult === 'object') {
            structureValue = structureResult[0] || structureResult.toString();
        }
        
        // 检查部首是否为空
        if (!radicalValue || radicalValue === '' || radicalValue === '[object Object]') {
            rowsToDelete.add(i);
            noRadicalCount++;
            continue;
        }
        
        // 检查是否为多音字 - CNChar格式: "(De|Dí|Dì)" 或 "De"
        if (pinyinWithTone.includes('(') && pinyinWithTone.includes('|')) {
            rowsToDelete.add(i);
            polyPhoneCount++;
            continue;
        }
        
        // 移除括号
        let pinyinStr = pinyinWithTone.replace(/[()]/g, '');
        
        // 将声调符号转换为数字
        const { pinyin: purePinyin, tone } = convertToneToNumber(pinyinStr);
        
        // 检查轻声
        if (tone === 0) {
            rowsToDelete.add(i);
            neutralToneCount++;
            continue;
        }
        
        // 检查是否匹配数字0-10
        const matchesNumber = numberPinyinTone.some(
            num => num.pinyin === purePinyin && num.tone === tone
        );
        
        if (matchesNumber) {
            rowsToDelete.add(i);
            numberMatchCount++;
            continue;
        }
        
        // 提取声母和韵母
        const { initial, final } = extractInitialAndFinal(purePinyin);
        
        // 添加新列数据
        data[i].push(
            strokeCount,
            purePinyin,
            radicalValue,
            structureValue,
            initial,
            final,
            tone
        );
        
    } catch (error) {
        console.log(`  行 ${i}: 字符 "${char}" - 处理出错: ${error.message}`);
        rowsToDelete.add(i);
        continue;
    }
    
    if (i % 500 === 0) {
        console.log(`  已处理 ${i} 行...`);
    }
}

console.log(`\n处理完成！准备删除不符合条件的行...`);
console.log(`\n删除统计:`);
console.log(`  多音字: ${polyPhoneCount}`);
console.log(`  轻声: ${neutralToneCount}`);
console.log(`  匹配数字: ${numberMatchCount}`);
console.log(`  无部首: ${noRadicalCount}`);
console.log(`  总删除: ${rowsToDelete.size}`);

// 创建过滤后的数据
const filteredData = data.filter((row, index) => !rowsToDelete.has(index));

console.log(`\n过滤前总行数（含表头）: ${data.length}`);
console.log(`过滤后总行数（含表头）: ${filteredData.length}`);
console.log(`保留的数据行数: ${filteredData.length - 1}`);

// 创建新的工作表
const newSheet = XLSX.utils.aoa_to_sheet(filteredData);
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Filtered Data');

// 保存文件
XLSX.writeFile(newWorkbook, OUTPUT_FILE);

console.log(`\n文件已保存: ${OUTPUT_FILE}`);
console.log('\n处理统计:');
console.log(`  原始数据行数: ${data.length - 1}`);
console.log(`  删除行数: ${rowsToDelete.size}`);
console.log(`  保留行数: ${filteredData.length - 1}`);
console.log(`  保留比例: ${((filteredData.length - 1) / (data.length - 1) * 100).toFixed(2)}%`);
