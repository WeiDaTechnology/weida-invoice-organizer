#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
从 PDF 发票中提取关键信息
支持中国增值税电子发票格式
"""

import pdfplumber
import re
import sys
import json
from pathlib import Path
from datetime import datetime

# 修复 Windows 控制台编码问题
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
# Mac/Linux 默认使用 UTF-8，无需特殊处理


def extract_invoice_info(pdf_path: str) -> dict:
    """
    从 PDF 发票中提取信息
    
    Returns:
        dict: {
            "invoice_code": str,      # 发票代码
            "invoice_number": str,    # 发票号码
            "date": str,              # 开票日期 (YYYY-MM-DD)
            "amount": float,          # 价税合计金额
            "buyer_name": str,        # 购买方名称
            "seller_name": str,       # 销售方名称
            "item_name": str,         # 项目名称/摘要
            "raw_text": str           # 原始提取文本 (用于调试)
        }
    """
    result = {
        "invoice_code": "",
        "invoice_number": "",
        "date": "",
        "amount": 0.0,
        "buyer_name": "",
        "seller_name": "",
        "item_name": "",
        "raw_text": "",
        "error": None
    }
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text() or ""
                full_text += text
            
            result["raw_text"] = full_text
            
            # 提取发票号码 (20 位数字)
            invoice_num_match = re.search(r'发票号码\s*[:：]?\s*(\d{20})', full_text)
            if invoice_num_match:
                result["invoice_number"] = invoice_num_match.group(1)
            
            # 提取发票代码 (20 位数字，如果有)
            invoice_code_match = re.search(r'发票代码\s*[:：]?\s*(\d{20})', full_text)
            if invoice_code_match:
                result["invoice_code"] = invoice_code_match.group(1)
            
            # 提取开票日期
            date_match = re.search(r'开票日期\s*[:：]?\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', full_text)
            if date_match:
                year, month, day = date_match.groups()
                result["date"] = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
            
            # 提取价税合计金额 (小写) - 优先匹配（小写）¥XXX.XX
            amount_match = re.search(r'（小写）\s*¥\s*([\d.]+)', full_text)
            if not amount_match:
                amount_match = re.search(r'\(小写\)\s*¥\s*([\d.]+)', full_text)
            if not amount_match:
                amount_match = re.search(r'价税合计.*?（小写）\s*¥\s*([\d.]+)', full_text)
            if not amount_match:
                amount_match = re.search(r'价税合计.*?¥\s*([\d.]+)', full_text)
            if amount_match:
                result["amount"] = float(amount_match.group(1))
            
            # 提取购买方名称
            buyer_match = re.search(r'购买方.*?名称\s*[:：]?\s*([^\n 售方]+)', full_text, re.DOTALL)
            if buyer_match:
                result["buyer_name"] = buyer_match.group(1).strip()
            
            # 更精确的购买方匹配
            buyer_match2 = re.search(r'买\s+名称\s*[:：]?\s*([^\n 售]+)', full_text)
            if buyer_match2:
                result["buyer_name"] = buyer_match2.group(1).strip()
            
            # 提取销售方名称
            seller_match = re.search(r'销售方.*?名称\s*[:：]?\s*([^\n]+)', full_text, re.DOTALL)
            if seller_match:
                result["seller_name"] = seller_match.group(1).strip()
            
            # 更精确的销售方匹配
            seller_match2 = re.search(r'售\s+名称\s*[:：]?\s*([^\n]+)', full_text)
            if seller_match2:
                result["seller_name"] = seller_match2.group(1).strip()
            
            # 提取项目名称/摘要 (第一行项目)
            # 匹配 *xxx* 格式，取第二个*后的内容
            item_match = re.search(r'\*[^*]+\*\s*([^\n]+)', full_text)
            if item_match:
                # 取第一个词作为摘要（如"通行费"）
                item_text = item_match.group(1).strip()
                # 只取第一个空格前的内容
                item_name = item_text.split()[0] if item_text else ""
                result["item_name"] = item_name
            
            # 如果没有提取到摘要，用销售方名称作为备选
            if not result["item_name"] and result["seller_name"]:
                # 取销售方名称的关键词（去掉"有限公司"等）
                name = result["seller_name"]
                name = re.sub(r'(有限公司 | 有限责任公司 | 公司)', '', name)
                result["item_name"] = name[:10]  # 限制长度
            
    except Exception as e:
        result["error"] = str(e)
    
    return result


def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_invoice_info.py <pdf_path>")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    result = extract_invoice_info(pdf_path)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
