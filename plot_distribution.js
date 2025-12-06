const XLSX = require('xlsx');
const plotly = require('plotly');
const fs = require('fs');
const path = require('path');

// 文件路径
const SUBTLEX_FILE = path.join(__dirname, 'SUBTLEX_List.xlsx');

console.log('开始读取Excel文件...\n');

// 读取Excel文件
const workbook = XLSX.readFile(SUBTLEX_FILE);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

console.log('表头:', data[0]);
console.log(`总行数: ${data.length}\n`);

// 提取C列（W/million，索引2）的数据
const wmillionData = [];

for (let i = 1; i < data.length; i++) {
    const value = data[i][2]; // C列
    
    // 只添加有效的数值
    if (value !== undefined && value !== null && value !== '' && !isNaN(value)) {
        wmillionData.push(Number(value));
    }
}

console.log(`C列 (W/million) 有效数据: ${wmillionData.length} 个`);

// 计算统计信息
const sorted = [...wmillionData].sort((a, b) => a - b);
const min = sorted[0];
const max = sorted[sorted.length - 1];
const mean = wmillionData.reduce((a, b) => a + b, 0) / wmillionData.length;
const median = sorted[Math.floor(sorted.length / 2)];

console.log(`\n统计信息:`);
console.log(`  最小值: ${min.toFixed(4)}`);
console.log(`  最大值: ${max.toFixed(4)}`);
console.log(`  均值: ${mean.toFixed(4)}`);
console.log(`  中位数: ${median.toFixed(4)}`);

// 生成HTML页面显示分布图
const htmlContent = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>W/million 分布曲线</title>
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 {
            color: #333;
            text-align: center;
        }
        .stats {
            background-color: #f9f9f9;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
        }
        .stats table {
            width: 100%;
            border-collapse: collapse;
        }
        .stats td {
            padding: 8px;
            border-bottom: 1px solid #ddd;
        }
        .stats td:first-child {
            font-weight: bold;
            width: 150px;
        }
        .chart {
            margin: 20px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>SUBTLEX_List.xlsx - W/million 分布曲线</h1>
        
        <div class="stats">
            <h3>统计信息</h3>
            <table>
                <tr>
                    <td>样本数量</td>
                    <td>${wmillionData.length}</td>
                </tr>
                <tr>
                    <td>最小值</td>
                    <td>${min.toFixed(4)}</td>
                </tr>
                <tr>
                    <td>最大值</td>
                    <td>${max.toFixed(4)}</td>
                </tr>
                <tr>
                    <td>均值</td>
                    <td>${mean.toFixed(4)}</td>
                </tr>
                <tr>
                    <td>中位数</td>
                    <td>${median.toFixed(4)}</td>
                </tr>
            </table>
        </div>
        
        <div id="histogram" class="chart"></div>
        <div id="density" class="chart"></div>
        <div id="boxplot" class="chart"></div>
    </div>
    
    <script>
        const data = ${JSON.stringify(wmillionData)};
        
        // 1. 直方图
        const histogramTrace = {
            x: data,
            type: 'histogram',
            nbinsx: 50,
            marker: {
                color: 'rgba(100, 150, 250, 0.7)',
                line: {
                    color: 'rgba(100, 150, 250, 1)',
                    width: 1
                }
            },
            name: 'W/million'
        };
        
        const histogramLayout = {
            title: '直方图分布',
            xaxis: { title: 'W/million' },
            yaxis: { title: '频数' },
            bargap: 0.05,
            height: 500
        };
        
        Plotly.newPlot('histogram', [histogramTrace], histogramLayout, {responsive: true});
        
        // 2. 核密度估计曲线（使用小箱子近似）
        const densityTrace = {
            x: data,
            type: 'histogram',
            histnorm: 'probability density',
            nbinsx: 100,
            marker: {
                color: 'rgba(250, 100, 100, 0.5)',
                line: {
                    color: 'rgba(250, 100, 100, 0.8)',
                    width: 0.5
                }
            },
            name: '密度'
        };
        
        const densityLayout = {
            title: '密度分布曲线',
            xaxis: { title: 'W/million' },
            yaxis: { title: '概率密度' },
            bargap: 0.01,
            height: 500
        };
        
        Plotly.newPlot('density', [densityTrace], densityLayout, {responsive: true});
        
        // 3. 箱线图
        const boxTrace = {
            y: data,
            type: 'box',
            name: 'W/million',
            marker: {
                color: 'rgba(100, 200, 150, 0.7)'
            },
            boxmean: 'sd'
        };
        
        const boxLayout = {
            title: '箱线图',
            yaxis: { title: 'W/million' },
            height: 500
        };
        
        Plotly.newPlot('boxplot', [boxTrace], boxLayout, {responsive: true});
    </script>
</body>
</html>`;

// 保存HTML文件
const htmlFilePath = path.join(__dirname, 'w_million_distribution.html');
fs.writeFileSync(htmlFilePath, htmlContent, 'utf8');

console.log(`\n分布曲线图已生成！`);
console.log(`文件路径: ${htmlFilePath}`);
console.log(`\n请在浏览器中打开该HTML文件查看分布曲线图。`);
