# Weida Invoice Organizer

微搭信息科技 - 发票自动整理工具

自动整理 PDF 发票，提取发票信息（日期/金额/发票号/商家），复制到新文件夹并重命名，导出到 Excel 报销清单模板。

## ✨ 功能特点

- ✅ **自动发现** - 扫描文件夹中的所有 PDF 发票文件
- ✅ **信息提取** - 从中国增值税电子发票中提取关键信息：
  - 发票号码
  - 开票日期
  - 价税合计金额
  - 销售方名称
  - 项目名称/摘要
- ✅ **智能重命名** - 按 `金额_摘要_发票号码.pdf` 格式重命名
- ✅ **原文件保护** - 原文件保持不变，复制到新文件夹
- ✅ **Excel 导出** - 自动生成费用报销清单明细表
- ✅ **内置模板** - 无需每次指定 Excel 模板

## 📦 安装

### 方式 1：克隆仓库

```bash
git clone https://github.com/WeiDaTechnology/weida-invoice-organizer.git
cd weida-invoice-organizer
```

### 方式 2：OpenClaw 技能安装（推荐）

如果你使用 OpenClaw，可以一键安装此技能：

```bash
# 在 OpenClaw 中执行
openclaw skills install https://github.com/WeiDaTechnology/weida-invoice-organizer.git
```

## 🔧 依赖

```bash
pip install pdfplumber openpyxl
```

## 🚀 快速开始

### 基础用法

```bash
# 整理文件夹中的发票（默认输出到 已整理 文件夹）
python scripts/organize_invoices.py "C:/path/to/invoices"

# 指定输出文件夹名称
python scripts/organize_invoices.py "C:/path/to/invoices" "已报销"
```

### 示例

```bash
# 整理 D 盘发票文件夹
python scripts/organize_invoices.py "D:/发票/2026 年 2 月"

# 输出到指定文件夹
python scripts/organize_invoices.py "D:/发票/2026 年 2 月" "2 月已报销"
```

## 📁 输出说明

运行后会在源文件夹下创建输出文件夹（默认名为 `已整理`），包含：

1. **重命名后的 PDF 发票** - 格式：`金额_摘要_发票号码.pdf`
   - 例如：`100_00_通行费_12345678901234567890.pdf`

2. **费用报销清单明细表.xlsx** - Excel 报销清单，包含：
   - 编号
   - 时间
   - 用途（销售方）
   - 金额
   - 责任人（留空）
   - 发票号
   - 摘要
   - 合计行（自动公式计算）

## 📋 发票信息提取说明

本工具支持中国增值税电子发票格式，可自动提取以下字段：

| 字段 | 说明 | 示例 |
|------|------|------|
| 发票号码 | 20 位数字 | 12345678901234567890 |
| 发票代码 | 20 位数字（如有） | 011002300111 |
| 开票日期 | YYYY-MM-DD 格式 | 2026-02-25 |
| 金额 | 价税合计（小写） | 100.00 |
| 销售方名称 | 开票方公司名称 | XX 科技有限公司 |
| 项目名称 | 发票明细第一项 | *通行费* |

## ⚠️ 注意事项

1. **PDF 格式**：仅支持中国增值税电子发票 PDF 格式
2. **原文件保护**：工具不会修改原始 PDF 文件，仅复制并重命名
3. **Excel 模板**：使用内置模板 `templates/费用报销清单明细表-demo.xlsx`
4. **失败处理**：提取失败的发票会记录在控制台输出中

## 🛠️ 开发

### 项目结构

```
weida-invoice-organizer/
├── scripts/
│   ├── organize_invoices.py      # 主脚本
│   └── extract_invoice_info.py   # 发票信息提取
├── templates/
│   └── 费用报销清单明细表-demo.xlsx  # Excel 模板
├── references/
│   └── invoice-formats.md        # 发票格式参考
├── SKILL.md                       # OpenClaw 技能定义
├── README.md                      # 本文件
└── .gitignore                     # Git 忽略规则
```

### 单独测试发票提取

```bash
python scripts/extract_invoice_info.py "path/to/invoice.pdf"
```

## 📄 许可证

MIT License

## 👥 作者

微搭信息科技 (WeiDa Technology)

## 🐛 问题反馈

如有问题或建议，请在 GitHub 仓库提交 Issue。
