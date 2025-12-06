const cnchar = require('cnchar');
const poly = require('cnchar-poly');
const order = require('cnchar-order');
const trad = require('cnchar-trad');
const radical = require('cnchar-radical');
const info = require('cnchar-info');

cnchar.use(poly, order, trad, radical, info);

// 测试几个常见字
const testChars = ['的', '我', '你', '是', '了', '一', '二', '三', '啊', '呢', '字', '汉'];

console.log('测试CNChar功能:\n');

testChars.forEach(char => {
    console.log(`\n字符: ${char}`);
    const spell1 = cnchar.spell(char);
    const spell2 = cnchar.spell(char, 'tone');
    const spell3 = cnchar.spell(char, 'tone', 'poly');
    
    console.log(`  拼音(默认): "${spell1}"`);
    console.log(`  拼音(tone): "${spell2}"`);
    console.log(`  拼音(tone+poly): "${spell3}"`);
    console.log(`  笔画数: ${cnchar.stroke(char)}`);
    console.log(`  部首: ${JSON.stringify(cnchar.radical(char))}`);
    console.log(`  字形结构: ${JSON.stringify(cnchar.info(char, 'struct'))}`);
    
    // 检查声调
    const toneMatch = spell2.match(/[1-4]/);
    console.log(`  提取的声调: ${toneMatch ? toneMatch[0] : '无'}`);
});
