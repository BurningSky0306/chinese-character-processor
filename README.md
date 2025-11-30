# 汉字属性处理工具

用于心理学实验刺激材料准备的汉字属性处理工具。从原始字频数据集一步生成符合实验要求的筛选结果。

## 🎯 功能特点

- **一步到位**：从 original 数据集直接生成 final 结果
- **自动提取**：汉字的拼音、笔画、部首、结构、声母、韵母等属性
- **智能筛选**：
  - 基于频率阈值过滤
  - 自动识别并排除多音字
  - 排除轻声字
  - 排除与数字0-10同音的汉字
- **基于权威库**：使用 [cnchar](https://github.com/theajack/cnchar) 汉字库

## 📦 安装

```bash
npm install
```

## 🚀 使用方法

### 方式一：自动检测文件（推荐）

```bash
npm start
```

脚本会自动查找目录中以 `original` 开头的 `.xlsx` 文件作为输入，输出文件名为 `filtered_Chinese character list.xlsx`。

### 方式二：指定输入文件

```bash
npm start mydata.xlsx
```

使用指定的输入文件，输出文件名自动生成为 `filtered_mydata.xlsx`。

### 方式三：完全自定义

```bash
npm start input.xlsx output.xlsx
```

同时指定输入和输出文件名。

**注意**：在 Windows PowerShell 中使用自定义参数时，需要使用 `--` 分隔符：

```bash
npm start -- mydata.xlsx result.xlsx
```

## 📂 数据集说明

### 输入文件
默认情况下，脚本会自动查找以 `original` 开头的 `.xlsx` 文件。示例：
- `original_Chinese character list from 2.5 billion words corpus ordered by frequency.xlsx`
  - 包含 14,976 个汉字
  - 按字频排序
  - 来源：25亿字语料库

您也可以使用命令行参数指定任意文件名（详见"使用方法"部分）。

### 输出文件
默认输出文件名为 `filtered_Chinese character list.xlsx`，也可以通过命令行参数自定义。

输出文件特点：
- 经过所有筛选规则处理后的最终结果
- 包含原始5列 + 新增7列共12列数据

### 数据集类型
1. **original_**: 原始数据集（字频统计）
2. **filtered_**: 客观筛选后的数据集（本脚本输出）
3. **distribution_**: 主观筛选后用于分发的数据集

## 📊 数据列说明

### 原始列（1-5）
1. **serial number**: 序号
2. **character**: 汉字本身
3. **token**: 在语料中的出现频数
4. **frequency(per million)**: 每百万字中的出现次数
5. **total coverage rate(%)**: 累积覆盖率

### 新增列（6-12）
6. **stroke count**: 笔画数
7. **pinyin**: 拼音（含声调）
8. **radical**: 部首
9. **structure**: 字形结构（如"左右结构"、"上下结构"）
10. **Initial consonant**: 声母
11. **final vowels**: 韵母
12. **tone**: 声调（1/2/3/4）

## 🔍 筛选规则

脚本会自动应用以下筛选规则：

1. ✅ **频率过滤**：保留 frequency(per million) > 0.98 的汉字
2. ✅ **笔画过滤**：排除笔画数为 0 的字符
3. ✅ **多音字过滤**：排除所有多音字
4. ✅ **轻声过滤**：排除轻声字（tone = 0）
5. ✅ **数字同音字过滤**：排除与数字 0-10 拼音声调完全相同的汉字
   - 零(líng) 一(yī) 二(èr) 三(sān) 四(sì) 五(wǔ)
   - 六(liù) 七(qī) 八(bā) 九(jiǔ) 十(shí)

## 🛠 技术栈

- **Node.js**: JavaScript 运行环境
- **cnchar**: 汉字拼音笔画库
  - cnchar-poly: 多音字支持
  - cnchar-radical: 部首查询
  - cnchar-order: 笔画顺序
- **xlsx**: Excel 文件处理

## 📈 典型处理结果

- 原始数据：~15,000 汉字
- 经过筛选：~3,000-3,500 汉字
- 保留比例：约 20-25%

## 🎓 使用场景

适用于以下研究领域的实验材料准备：
- 心理学（认知心理学、发展心理学）
- 语言学（汉语研究、语言习得）
- 认知科学（语言认知、阅读研究）
- 教育学（识字教学、语文教育）

## 📝 许可证

MIT

## 👨‍💻 作者

用于心理学实验材料准备

## 🔗 相关链接

- [cnchar 官方文档](https://theajack.github.io/cnchar/)
- [cnchar GitHub](https://github.com/theajack/cnchar)
