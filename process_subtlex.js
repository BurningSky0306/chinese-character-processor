/**
 * SUBTLEX 汉字数据处理脚本
 * 参考 main.js 的逻辑处理 SUBTLEX_List.xlsx
 */

const XLSX = require('xlsx');
const cnchar = require('cnchar');
const poly = require('cnchar-poly');
const radical = require('cnchar-radical');
const order = require('cnchar-order');
const path = require('path');

// 加载 cnchar 插件
cnchar.use(poly, radical, order);

console.log('==========================================');
console.log('   SUBTLEX 汉字数据处理脚本');
console.log('==========================================\n');

// 文件路径
const INPUT_FILE = path.join(__dirname, 'SUBTLEX_List.xlsx');
const OUTPUT_FILE = path.join(__dirname, 'filtered_SUBTLEX_List.xlsx');

// 数字拼音对照表 (0-10)
const NUMBER_PINYIN_MAP = [
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

// 统计信息
const stats = {
    totalRows: 0,
    excludedByPolyphone: 0,
    excludedByLightTone: 0,
    excludedByNumberPinyin: 0,
    excludedByNoRadical: 0,
    excludedByOther: 0,
    kept: 0
};

/**
 * 检查是否为数字同音字
 */
function isNumberPinyin(pinyin, tone) {
    return NUMBER_PINYIN_MAP.some(num => 
        num.pinyin === pinyin && num.tone === tone
    );
}

/**
 * 提取汉字的所有属性
 */
function extractCharacterProperties(character) {
    const result = {
        strokeCount: 0,
        pinyin: '',
        radical: '',
        structure: '',
        initial: '',
        final: '',
        tone: 0,
        isValid: false,
        reason: ''
    };

    try {
        // 1. 笔画数
        result.strokeCount = cnchar.stroke(character);

        // 2. 拼音 - 检查是否为多音字
        const pinyinResult = cnchar.spell(character, 'tone');
        
        // 检查是否为多音字
        if (cnchar.isPolyWord(character)) {
            result.reason = '多音字';
            return result;
        }
        
        result.pinyin = pinyinResult;

        // 3. 部首
        try {
            const radicalResult = cnchar.radical(character);
            if (Array.isArray(radicalResult) && radicalResult.length > 0) {
                result.radical = radicalResult[0].radical || '';
            }
        } catch (e) {
            result.radical = '';
        }

        // 检查部首是否为空（字符不在库中）
        if (!result.radical || result.radical === '') {
            result.reason = '无部首';
            return result;
        }

        // 4. 结构
        try {
            const radicalResult = cnchar.radical(character);
            if (Array.isArray(radicalResult) && radicalResult.length > 0) {
                result.structure = radicalResult[0].struct || '';
            }
        } catch (e) {
            result.structure = '';
        }

        // 5. 声母、韵母和声调
        try {
            const spellInfoResult = cnchar.spellInfo(result.pinyin);
            result.initial = spellInfoResult.initial || '';
            result.final = spellInfoResult.final || '';
            result.tone = spellInfoResult.tone;
        } catch (e) {
            result.initial = '';
            result.final = '';
            result.tone = 0;
        }

        // 检查是否为轻声（tone = 0）
        if (result.tone === 0) {
            result.reason = '轻声字';
            return result;
        }

        // 检查是否为数字同音字
        const pinyinWithoutTone = result.pinyin.toLowerCase().replace(/[āáǎàēéěèīíǐìōóǒòūúǔùǖǘǚǜ]/g, match => {
            const toneMap = {
                'ā': 'a', 'á': 'a', 'ǎ': 'a', 'à': 'a',
                'ē': 'e', 'é': 'e', 'ě': 'e', 'è': 'e',
                'ī': 'i', 'í': 'i', 'ǐ': 'i', 'ì': 'i',
                'ō': 'o', 'ó': 'o', 'ǒ': 'o', 'ò': 'o',
                'ū': 'u', 'ú': 'u', 'ǔ': 'u', 'ù': 'u',
                'ǖ': 'v', 'ǘ': 'v', 'ǚ': 'v', 'ǜ': 'v'
            };
            return toneMap[match] || match;
        });

        if (isNumberPinyin(pinyinWithoutTone, result.tone)) {
            result.reason = '数字同音字';
            return result;
        }

        result.isValid = true;
        return result;

    } catch (error) {
        result.reason = `处理错误: ${error.message}`;
        return result;
    }
}

/**
 * 主处理函数
 */
function processCharacters() {
    console.log(`正在读取文件: ${INPUT_FILE}...`);

    // 读取 Excel 文件
    const workbook = XLSX.readFile(INPUT_FILE);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    console.log(`成功读取 ${data.length} 行数据\n`);
    console.log('正在处理汉字...');

    // 存储过滤后的数据
    const filteredData = [];

    // 添加新表头
    const headerRow = data[0];
    const newHeaders = [
        ...headerRow,
        'stroke count',
        'pinyin',
        'radical',
        'structure',
        'Initial consonant',
        'final vowels',
        'tone'
    ];
    filteredData.push(newHeaders);

    // 处理数据行(从第二行开始)
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        if (!row || row.length === 0) {
            continue;
        }

        stats.totalRows++;

        // 获取字符（第一列 Word）
        const character = row[0];

        // 检查是否为单个汉字
        if (!character || typeof character !== 'string' || character.length !== 1) {
            stats.excludedByOther++;
            continue;
        }

        // 提取汉字属性
        const props = extractCharacterProperties(character);

        if (!props.isValid) {
            // 根据失败原因统计
            if (props.reason === '多音字') {
                stats.excludedByPolyphone++;
            } else if (props.reason === '轻声字') {
                stats.excludedByLightTone++;
            } else if (props.reason === '数字同音字') {
                stats.excludedByNumberPinyin++;
            } else if (props.reason === '无部首') {
                stats.excludedByNoRadical++;
            } else {
                stats.excludedByOther++;
            }
            continue;
        }

        // 通过所有检查，保留这一行并添加新列
        const newRow = [
            ...row,
            props.strokeCount,
            props.pinyin,
            props.radical,
            props.structure,
            props.initial,
            props.final,
            props.tone
        ];
        filteredData.push(newRow);
        stats.kept++;

        // 显示进度
        if (stats.totalRows % 1000 === 0) {
            console.log(`已处理 ${stats.totalRows} 行...`);
        }
    }

    console.log(`\n处理完成!`);
    console.log(`==========================================`);
    console.log(`统计信息:`);
    console.log(`  总处理行数: ${stats.totalRows}`);
    console.log(`  排除多音字: ${stats.excludedByPolyphone}`);
    console.log(`  排除轻声字: ${stats.excludedByLightTone}`);
    console.log(`  排除数字同音字: ${stats.excludedByNumberPinyin}`);
    console.log(`  排除无部首: ${stats.excludedByNoRadical}`);
    console.log(`  其他错误: ${stats.excludedByOther}`);
    console.log(`  保留的行数: ${stats.kept}`);
    console.log(`  保留比例: ${(stats.kept / stats.totalRows * 100).toFixed(2)}%`);
    console.log(`==========================================\n`);

    // 创建新的工作表
    const newWorksheet = XLSX.utils.aoa_to_sheet(filteredData);

    // 创建新的工作簿
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Filtered Data');

    // 写入文件
    console.log(`正在保存到文件: ${OUTPUT_FILE}...`);
    XLSX.writeFile(newWorkbook, OUTPUT_FILE);

    console.log('\n==========================================');
    console.log('   处理完成!');
    console.log(`   输出文件: ${OUTPUT_FILE}`);
    console.log('==========================================\n');
    console.log('新增的列:');
    console.log('  - stroke count (笔画数)');
    console.log('  - pinyin (拼音，含声调)');
    console.log('  - radical (部首)');
    console.log('  - structure (结构)');
    console.log('  - Initial consonant (声母)');
    console.log('  - final vowels (韵母)');
    console.log('  - tone (声调: 1/2/3/4)');
    console.log('\n筛选规则:');
    console.log('  1. 排除多音字');
    console.log('  2. 排除轻声字');
    console.log('  3. 排除与数字0-10拼音声调相同的汉字');
    console.log('  4. 排除无部首的字符（不在CNChar库中）');
}

// 执行主函数
processCharacters();
