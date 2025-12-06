const XLSX = require('xlsx');
const jStat = require('jstat');
const path = require('path');

// 文件路径
const CHARACTER_FILE = path.join(__dirname, 'CHARACTER_List.xlsx');

console.log('开始读取Excel文件...\n');

// 读取Excel文件
const workbook = XLSX.readFile(CHARACTER_FILE);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

console.log('表头:', data[0]);
console.log(`总行数: ${data.length}\n`);

// 提取D列（frequency per million，索引3）和N列（W/million，索引13）
const frequencyData = [];
const wmillionData = [];

for (let i = 1; i < data.length; i++) {
    const freq = data[i][3]; // D列
    const wmil = data[i][13]; // N列
    
    // 只添加有效的数值
    if (freq !== undefined && freq !== null && freq !== '' && !isNaN(freq)) {
        frequencyData.push(Number(freq));
    }
    if (wmil !== undefined && wmil !== null && wmil !== '' && !isNaN(wmil)) {
        wmillionData.push(Number(wmil));
    }
}

console.log(`D列 (frequency per million) 有效数据: ${frequencyData.length} 个`);
console.log(`N列 (W/million) 有效数据: ${wmillionData.length} 个\n`);

// 确保两组数据长度相同（配对样本）
const minLength = Math.min(frequencyData.length, wmillionData.length);
const data1 = frequencyData.slice(0, minLength);
const data2 = wmillionData.slice(0, minLength);

console.log(`使用配对样本数量: ${minLength}\n`);

// ===== 1. 配对样本t检验 =====
console.log('========== 配对样本t检验 ==========');

// 计算差值
const differences = data1.map((val, i) => val - data2[i]);

// 计算均值
const mean1 = jStat.mean(data1);
const mean2 = jStat.mean(data2);
const meanDiff = jStat.mean(differences);

// 计算标准差
const sd1 = jStat.stdev(data1, true);
const sd2 = jStat.stdev(data2, true);
const sdDiff = jStat.stdev(differences, true);

// 计算t值
const n = differences.length;
const sem = sdDiff / Math.sqrt(n);
const tValue = meanDiff / sem;

// 计算自由度
const df = n - 1;

// 计算双尾p值
const pValue = 2 * (1 - jStat.studentt.cdf(Math.abs(tValue), df));

// 计算Cohen's d效应量
const pooledSD = Math.sqrt((sd1 * sd1 + sd2 * sd2) / 2);
const cohensD = meanDiff / pooledSD;

// 计算95%置信区间
const tCritical = jStat.studentt.inv(0.975, df);
const marginOfError = tCritical * sem;
const ci95Lower = meanDiff - marginOfError;
const ci95Upper = meanDiff + marginOfError;

console.log(`样本数量 (n): ${n}`);
console.log(`D列均值: ${mean1.toFixed(4)}`);
console.log(`N列均值: ${mean2.toFixed(4)}`);
console.log(`均值差: ${meanDiff.toFixed(4)}`);
console.log(`标准差差: ${sdDiff.toFixed(4)}`);
console.log(`t值: ${tValue.toFixed(4)}`);
console.log(`自由度 (df): ${df}`);
console.log(`p值: ${pValue.toExponential(4)}`);
console.log(`Cohen's d (效应量): ${cohensD.toFixed(4)}`);
console.log(`95% 置信区间: [${ci95Lower.toFixed(4)}, ${ci95Upper.toFixed(4)}]`);

// ===== 2. 置换检验 (Permutation Test) =====
console.log('\n========== 置换检验 ==========');

const permutations = 10000;
let countExtremeOrMore = 0;
const observedDiff = Math.abs(meanDiff);

console.log(`进行 ${permutations} 次置换...`);

for (let i = 0; i < permutations; i++) {
    // 随机置换差值的符号
    const permutedDiff = differences.map(d => Math.random() < 0.5 ? d : -d);
    const permutedMean = jStat.mean(permutedDiff);
    
    if (Math.abs(permutedMean) >= observedDiff) {
        countExtremeOrMore++;
    }
    
    if ((i + 1) % 2000 === 0) {
        console.log(`  已完成 ${i + 1} 次置换...`);
    }
}

const permutationPValue = countExtremeOrMore / permutations;

console.log(`观察到的均值差绝对值: ${observedDiff.toFixed(4)}`);
console.log(`极端或更极端的置换次数: ${countExtremeOrMore}`);
console.log(`置换检验 p值: ${permutationPValue.toFixed(4)}`);

// ===== 3. Bootstrap置信区间 =====
console.log('\n========== Bootstrap 置信区间 ==========');

const bootstrapSamples = 10000;
const bootstrapMeans = [];

console.log(`进行 ${bootstrapSamples} 次Bootstrap重采样...`);

for (let i = 0; i < bootstrapSamples; i++) {
    // 有放回的重采样
    const resampledDiff = [];
    for (let j = 0; j < n; j++) {
        const randomIndex = Math.floor(Math.random() * n);
        resampledDiff.push(differences[randomIndex]);
    }
    
    const bootstrapMean = jStat.mean(resampledDiff);
    bootstrapMeans.push(bootstrapMean);
    
    if ((i + 1) % 2000 === 0) {
        console.log(`  已完成 ${i + 1} 次重采样...`);
    }
}

// 对bootstrap均值排序
bootstrapMeans.sort((a, b) => a - b);

// 计算95%置信区间（百分位法）
const lowerIndex = Math.floor(bootstrapSamples * 0.025);
const upperIndex = Math.floor(bootstrapSamples * 0.975);
const bootstrapCI95Lower = bootstrapMeans[lowerIndex];
const bootstrapCI95Upper = bootstrapMeans[upperIndex];

console.log(`Bootstrap 95% 置信区间 (百分位法): [${bootstrapCI95Lower.toFixed(4)}, ${bootstrapCI95Upper.toFixed(4)}]`);

// 计算99%置信区间
const lower99Index = Math.floor(bootstrapSamples * 0.005);
const upper99Index = Math.floor(bootstrapSamples * 0.995);
const bootstrapCI99Lower = bootstrapMeans[lower99Index];
const bootstrapCI99Upper = bootstrapMeans[upper99Index];

console.log(`Bootstrap 99% 置信区间 (百分位法): [${bootstrapCI99Lower.toFixed(4)}, ${bootstrapCI99Upper.toFixed(4)}]`);

// ===== 汇总结果 =====
console.log('\n========== 结果汇总 ==========');
console.log('【t检验】');
console.log(`  t值: ${tValue.toFixed(4)}`);
console.log(`  p值: ${pValue.toExponential(4)}`);
console.log(`  Cohen's d (效应量): ${cohensD.toFixed(4)}`);
console.log(`  95% 置信区间: [${ci95Lower.toFixed(4)}, ${ci95Upper.toFixed(4)}]`);
console.log('\n【置换检验】');
console.log(`  p值: ${permutationPValue.toFixed(4)}`);
console.log('\n【Bootstrap】');
console.log(`  95% 置信区间: [${bootstrapCI95Lower.toFixed(4)}, ${bootstrapCI95Upper.toFixed(4)}]`);
console.log(`  99% 置信区间: [${bootstrapCI99Lower.toFixed(4)}, ${bootstrapCI99Upper.toFixed(4)}]`);

// 保存结果到文件
const results = {
    sampleSize: n,
    means: {
        frequencyPerMillion: mean1,
        wPerMillion: mean2,
        difference: meanDiff
    },
    tTest: {
        tValue: tValue,
        pValue: pValue,
        degreesOfFreedom: df,
        cohensD: cohensD,
        confidenceInterval95: {
            lower: ci95Lower,
            upper: ci95Upper
        }
    },
    permutationTest: {
        permutations: permutations,
        pValue: permutationPValue
    },
    bootstrap: {
        samples: bootstrapSamples,
        confidenceInterval95: {
            lower: bootstrapCI95Lower,
            upper: bootstrapCI95Upper
        },
        confidenceInterval99: {
            lower: bootstrapCI99Lower,
            upper: bootstrapCI99Upper
        }
    }
};

const fs = require('fs');
fs.writeFileSync(
    path.join(__dirname, 'statistical_results.json'),
    JSON.stringify(results, null, 2),
    'utf8'
);

console.log('\n结果已保存到 statistical_results.json');
