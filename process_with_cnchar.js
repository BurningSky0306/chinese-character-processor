const XLSX = require('xlsx');
const path = require('path');

// 导入cnchar及其插件
const cnchar = require('cnchar');
const poly = require('cnchar-poly'); // 多音字插件
const order = require('cnchar-order'); // 笔画顺序插件
const trad = require('cnchar-trad'); // 繁体字插件
const radical = require('cnchar-radical'); // 部首插件
const info = require('cnchar-info'); // 汉字信息插件

// 使用插件
cnchar.use(poly, order, trad, radical, info);

// 提取声母和韵母的辅助函数
function extractInitialAndFinal(pinyin) {
    // 去除数字声调，转为小写
    const pure = pinyin.replace(/[1-4]/g, '').toLowerCase();
    
    // 声母列表
    const initials = ['b', 'p', 'm', 'f', 'd', 't', 'n', 'l', 'g', 'k', 'h', 
                     'j', 'q', 'x', 'zh', 'ch', 'sh', 'r', 'z', 'c', 's', 'y', 'w'];
    
    let initial = '';
    let final = pure;
    
    // 检查双字母声母
    for (const init of ['zh', 'ch', 'sh']) {
        if (pure.startsWith(init)) {
            initial = init;
            final = pure.substring(init.length);
            return { initial, final };
        }
    }
    
    // 检查单字母声母
    for (const init of initials) {
        if (pure.startsWith(init)) {
            initial = init;
            final = pure.substring(init.length);
            break;
        }
    }
    
    // 零声母的韵母
    if (!initial) {
        final = pure;
    }
    
    return { initial, final };
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
console.log('新表头:', data[0]);

// 存储需要删除的行索引
const rowsToDelete = new Set();

// 数字0-10的拼音和声调对照
const numberPinyinTone = [
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

console.log('\n开始处理每个字符...');

// 处理每一行（从第2行开始，索引1）
for (let i = 1; i < data.length; i++) {
    const word = data[i][0]; // A列 - Word
    
    if (!word || typeof word !== 'string' || word.length !== 1) {
        // 不是单个汉字，标记删除
        rowsToDelete.add(i);
        continue;
    }
    
    const char = word;
    
    try {
        // 1. 获取笔画数
        const strokeCount = cnchar.stroke(char);
        
        // 2. 获取拼音（带声调）
        const pinyinWithTone = cnchar.spell(char, 'tone', 'poly');
        
        // 3. 获取部首
        const radical = cnchar.radical(char);
        
        // 4. 获取字形结构
        const structure = cnchar.info(char, 'struct');
        
        // 检查部首是否为空（字符不在库中）
        if (!radical || radical === '') {
            rowsToDelete.add(i);
            if (i % 500 === 0) {
                console.log(`  行 ${i}: 字符 "${char}" - 部首为空，删除`);
            }
            continue;
        }
        
        // 检查是否为多音字
        const pinyinArray = pinyinWithTone.split(',');
        if (pinyinArray.length > 1) {
            rowsToDelete.add(i);
            if (i % 500 === 0) {
                console.log(`  行 ${i}: 字符 "${char}" - 多音字，删除`);
            }
            continue;
        }
        
        const pinyin = pinyinArray[0];
        
        // 提取声调数字
        const toneMatch = pinyin.match(/[1-4]/);
        const tone = toneMatch ? parseInt(toneMatch[0]) : 0;
        
        // 检查是否为轻声（tone为0或没有声调）
        if (tone === 0) {
            rowsToDelete.add(i);
            if (i % 500 === 0) {
                console.log(`  行 ${i}: 字符 "${char}" - 轻声，删除`);
            }
            continue;
        }
        
        // 提取纯拼音（去除声调数字）
        const purePinyin = pinyin.replace(/[1-4]/g, '').toLowerCase();
        
        // 检查是否匹配数字0-10的拼音和声调
        const matchesNumber = numberPinyinTone.some(
            num => num.pinyin === purePinyin && num.tone === tone
        );
        
        if (matchesNumber) {
            rowsToDelete.add(i);
            if (i % 500 === 0) {
                console.log(`  行 ${i}: 字符 "${char}" - 匹配数字读音，删除`);
            }
            continue;
        }
        
        // 5. 获取声母
        const initialConsonant = cnchar.initial(char);
        
        // 6. 获取韵母
        const finalVowels = cnchar.final(char);
        
        // 7. 声调已经提取
        
        // 添加新列数据
        data[i].push(
            strokeCount,           // stroke count
            purePinyin,            // pinyin (不带声调)
            radical,               // radical
            structure,             // structure
            initialConsonant,      // Initial consonant
            finalVowels,           // final vowels
            tone                   // tone
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
console.log(`需要删除的行数: ${rowsToDelete.size}`);

// 创建过滤后的数据（保留未被标记删除的行）
const filteredData = data.filter((row, index) => !rowsToDelete.has(index));

console.log(`\n过滤前总行数（含表头）: ${data.length}`);
console.log(`过滤后总行数（含表头）: ${filteredData.length}`);
console.log(`保留的数据行数: ${filteredData.length - 1}`);

// 统计删除原因
let polyPhoneCount = 0;
let neutralToneCount = 0;
let numberMatchCount = 0;
let noRadicalCount = 0;

// 重新统计各类删除原因（可选，用于详细报告）
console.log('\n生成新的Excel文件...');

// 创建新的工作表
const newSheet = XLSX.utils.aoa_to_sheet(filteredData);

// 创建新的工作簿
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Filtered Data');

// 保存文件
XLSX.writeFile(newWorkbook, OUTPUT_FILE);

console.log(`\n文件已保存: ${OUTPUT_FILE}`);
console.log('\n处理统计:');
console.log(`  原始数据行数: ${data.length - 1}`);
console.log(`  删除行数: ${rowsToDelete.size}`);
console.log(`  保留行数: ${filteredData.length - 1}`);
console.log(`  删除比例: ${(rowsToDelete.size / (data.length - 1) * 100).toFixed(2)}%`);
