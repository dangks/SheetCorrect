# Excel 数据匹配与修改工具

一个基于浏览器的 Excel 数据处理工具，主要功能是读取Excel表格、进行两表数据匹配、比较两表差异，并就地修改数据及支持导出修正版本。   

## 功能特性

- 📊 支持多工作表的 Excel 文件
- 🔍 灵活的数据匹配规则
- ✏️ 单个/一键批量数据更新
- 📋 实时预览数据变更
- 🎨 直观的界面操作


## 使用说明

1. 选择源数据 Excel 文件（包含正确数据的文件）
2. 选择目标 Excel 文件（需要更新的文件）
3. 选择匹配字段（用于确定相同记录）
4. 选择需要更新的字段
5. 点击"开始数据匹配"进行比对
6. 检查修改建议，可单个或批量应用修改
7. 导出保留原格式的 Excel 文件


## 技术栈

- 原生 JavaScript
- SheetJS (XLSX) 用于 Excel 文件处理
- HTML5 / CSS3


## 安装和运行

1. 克隆仓库：
```bash
git clone [仓库地址]
```

2. 打开项目文件夹：
```bash
cd SheetCorrect
```

3. 在浏览器中打开 index.html 文件即可使用


## 项目结构

```
SheetCorrect/
├── index.html          # 主页面
├── src/
│   ├── js/
│   │   ├── main.js    # 主要业务逻辑
│   │   └── xlsx.full.min.js  # Excel处理库
│   └── css/
│       └── styles.css  # 样式文件
└── README.md          # 项目说明文档
```

## 注意事项

- 请确保上传的 Excel 文件格式正确（.xlsx 或 .xls）
- 建议先进行小范围测试，确认修改符合预期
- 导出时会自动备份，文件名会包含时间戳
- 支持日期、文本、数字等常见数据类型
- 不要直接修改正式文件，请先备份数据


## 开源致谢

本项目使用了以下开源项目，在此特别感谢：

- [SheetJS](https://github.com/SheetJS/sheetjs) - 优秀的 JavaScript Excel 文件处理库


## 许可证

[MIT License](LICENSE)


## ToDo

- 💾 保留原始 Excel 格式导出
- 📝 添加更多数据类型支持
- 🐞 支持更强大的数据处理
