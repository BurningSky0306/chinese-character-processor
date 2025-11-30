# 汉字属性处理工具

用于心理学实验刺激材料准备的汉字属性处理脚本集合。

## 功能特点

- 自动提取汉字的拼音、笔画、部首、结构、声母、韵母等属性
- 支持多音字识别和过滤
- 支持自定义筛选规则（笔画数、声调、数字同音字等）
- 基于 [cnchar](https://github.com/theajack/cnchar) 汉字库

## 安装

```bash
npm install
```

## 使用方法

### 1. 处理汉字属性
为汉字列表添加拼音、笔画、部首等属性：
```bash
npm start
```

### 2. 过滤笔画数为0和多音字
```bash
npm run filter
```

### 3. 高级过滤
基于 serial number 和数字同音字过滤：
```bash
npm run advanced-filter
```

### 4. 人工审核后处理
对人工审核过的列表进行处理：
```bash
npm run manual-filter
```

## 输入输出

### 输入文件
- `Chinese character list from 2.5 billion words corpus ordered by frequency.xlsx`

### 输出文件
- `Chinese character list with properties.xlsx` - 添加属性后的完整列表
- `Chinese character list filtered.xlsx` - 过滤后的列表
- `Chinese character list final.xlsx` - 最终筛选结果
- `Chinese character list final manual.xlsx` - 人工审核后的最终结果

## 添加的汉字属性

- **stroke count**: 笔画数
- **pinyin**: 拼音（含音调）
- **radical**: 部首
- **structure**: 字形结构
- **Initial consonant**: 声母
- **final vowels**: 韵母
- **tone**: 声调（1/2/3/4/5，其中5表示轻声）

## 筛选规则

- 排除笔画数为 0 的字符
- 排除多音字
- 排除与数字 0-10 拼音声调相同的汉字
- 支持按字频排序筛选

## 依赖库

- [cnchar](https://github.com/theajack/cnchar) - 功能全面的汉字拼音笔画 js 库
- [xlsx](https://www.npmjs.com/package/xlsx) - Excel 文件处理

## 使用场景

适用于心理学、语言学、认知科学等领域的实验材料准备。

## 许可证

MIT

## 作者

制作于心理学实验材料准备
