---
name: weida-invoice-organizer
description: 自动整理 PDF 发票的技能。支持扫描文件夹发现发票、提取发票信息（日期/金额/发票号/商家）、复制到新文件夹并重命名、导出到 Excel 模板。适用于中国增值税电子发票的批量处理场景。
---

# 发票整理技能

自动整理 PDF 发票，提取信息并复制到新文件夹重命名，导出到 Excel 报销清单。

**特点：**
- ✅ 原文件保持不变，复制到新文件夹
- ✅ 默认输出到 `已整理` 子文件夹
- ✅ 内置 Excel 模板，无需每次指定

## 快速开始

```bash
# 基础用法 - 整理文件夹中的发票（默认输出到 已整理 文件夹）
python scripts/organize_invoices.py "C:/path/to/invoices"

# 指定输出文件夹名称
python scripts/organize_invoices.py "C:/path/to/invoices" "已报销"
```

## 功能

1. **自动发现** - 扫描文件夹中的所有 PDF 文件
2. **信息提取** - 从发票中提取：
   - 发票号码
   - 开票日期
   - 金额（价税合计）
   - 销售方名称
   - 项目名称/摘要
3. **复制并重命名** - 复制到新文件夹，按规则重命名：`金额_摘要_发票号码.pdf`
4. **Excel 导出** - 使用内置模板填入报销清单（含统计行）

## 重命名规则

```
{金额}_{摘要}_{发票号码}.pdf

示例：0_47_通行费_26317907100200441189.pdf
```

## Excel 模板格式

模板应包含以下列（按顺序）：

| 列 | 字段 | 说明 |
|---|---|---|
| 1 | 编号 | 自动序号 |
| 2 | 时间 | 开票日期 |
| 3 | 用途（详细用途） | 销售方名称 |
| 4 | 金额 | 价税合计 |
| 5 | 责任人 | 留空 |
| 6 | 发票号 | 发票号码 |
| 7 | 摘要 | 项目名称 |

**统计行**：最后一行自动添加合计，格式为 `合计：总金额`

内置模板：`templates/费用报销清单明细表-demo.xlsx`

## 脚本说明

### organize_invoices.py

主处理脚本，orchestrate 整个流程。

**参数：**
- `folder_path` (必需): 发票 PDF 所在文件夹
- `output_folder_name` (可选): 输出文件夹名称（默认"已整理"）

**输出：**
- 在源文件夹下创建输出文件夹（如 `已整理`）
- 复制并重命名 PDF 文件到输出文件夹
- `费用报销清单明细表.xlsx` (自动生成)

### extract_invoice_info.py

底层发票信息提取工具。

**用法：**
```bash
python extract_invoice_info.py invoice.pdf
```

**返回 JSON：**
```json
{
  "invoice_code": "发票代码",
  "invoice_number": "发票号码",
  "date": "YYYY-MM-DD",
  "amount": 0.00,
  "buyer_name": "购买方",
  "seller_name": "销售方",
  "item_name": "项目名称",
  "raw_text": "原始文本"
}
```

## 依赖

```bash
pip install pdfplumber openpyxl
```

## 使用示例

### 示例 1: 整理待报销文件夹

```bash
cd skills/invoice-organizer
python scripts/organize_invoices.py "C:/Users/zhiliang.liu/Desktop/待报销/小金库"
```

### 示例 2: 完整流程（带 Excel 导出）

```bash
python scripts/organize_invoices.py \
  "C:/Users/zhiliang.liu/Desktop/待报销/小金库" \
  "C:/Users/zhiliang.liu/Desktop/待报销/小金库/费用报销清单明细表-demo.xlsx" \
  "C:/Users/zhiliang.liu/Desktop/待报销/小金库/已整理"
```

### 示例 3: 仅提取单张发票信息

```bash
python scripts/extract_invoice_info.py "C:/Users/zhiliang.liu/Desktop/待报销/小金库/0.47_toll.pdf"
```

## 支持的发票类型

- ✅ 增值税电子普通发票
- ✅ 电子发票（普通发票）
- ✅ 通行费电子发票
- ⚠️ 定额发票（需要 OCR，暂不支持）
- ⚠️ 手写发票（需要 OCR，暂不支持）

## 故障排除

### 问题：提取不到发票信息

**原因：** PDF 是扫描件（图片格式），不是文字 PDF

**解决：** 需要 OCR 支持。可以：
1. 使用 Adobe Acrobat 将扫描件转为文字 PDF
2. 或添加 OCR 模块（如 PaddleOCR）

### 问题：金额提取错误

**原因：** 发票格式不标准

**解决：** 检查 `extract_invoice_info.py` 中的正则表达式，根据实际发票格式调整。

### 问题：Excel 导出失败

**原因：** 模板格式不匹配

**解决：** 确保模板第一行是表头，从第二行开始是数据。参考 `费用报销清单明细表-demo.xlsx`。

## 相关文件

- [`scripts/organize_invoices.py`](scripts/organize_invoices.py) - 主处理脚本
- [`scripts/extract_invoice_info.py`](scripts/extract_invoice_info.py) - 发票信息提取
- [`references/invoice-formats.md`](references/invoice-formats.md) - 发票格式参考

## 注意事项

1. **备份原文件** - 处理前建议备份原始 PDF
2. **检查提取结果** - 首次使用建议先测试几张发票
3. **责任人字段** - 默认留空，可自行填写
4. **字符限制** - 摘要字段过长会被截断到 10 字符
5. **统计行** - Excel 最后一行自动添加金额合计
