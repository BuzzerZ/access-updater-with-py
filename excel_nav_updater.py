#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_nav_updater.py - 自动更新基金股票净值
修复版本：确保函数正确定义
"""

import pandas as pd
import requests
import re
import os
import sys
import time
from datetime import datetime
from typing import Optional

# ==================== 工具函数定义 ====================

def is_fund_code(code):
    """判断是否为基金代码"""
    if not code or pd.isna(code):
        return False
    code_str = str(int(code)) if isinstance(code, float) else str(code)
    if not code_str.isdigit() or len(code_str) != 6:
        return False
    return code_str[:2] in ['00', '01', '02', '15', '16', '18', '50', '51']

def is_stock_code(code):
    """判断是否为A股股票代码"""
    if not code or pd.isna(code):
        return False
    code_str = str(int(code)) if isinstance(code, float) else str(code)
    if not code_str.isdigit() or len(code_str) != 6:
        return False
    return code_str[:2] in ['60', '68', '30', '00']

def get_fund_nav(fund_code: str) -> Optional[float]:
    """获取基金净值"""
    try:
        code_str = str(int(fund_code)) if isinstance(fund_code, float) else str(fund_code)
        code_str = code_str.zfill(6)
        
        print(f"正在获取基金 {code_str} 净值...")
        
        url = "http://api.fund.eastmoney.com/f10/lsjz"
        params = {"fundCode": code_str, "pageIndex": 1, "pageSize": 1}
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Referer": f"http://fundf10.eastmoney.com/jjjz_{code_str}.html"
        }
        
        response = requests.get(url, params=params, headers=headers, timeout=10)
        data = response.json()
        
        if data and "Data" in data and data["Data"]:
            lst = data["Data"].get("LSJZList") or []
            if lst:
                nav_str = lst[0].get("DWJZ")
                if nav_str and nav_str != "":
                    nav_value = float(nav_str)
                    print(f"✅ 基金 {code_str} 净值: {nav_value}")
                    return nav_value
        print(f"❌ 基金 {code_str} 未获取到净值数据")
    except Exception as e:
        print(f"❌ 获取基金 {fund_code} 净值失败: {e}")
    
    return None

def get_stock_price(stock_code: str) -> Optional[float]:
    """获取股票价格"""
    try:
        code_str = str(int(stock_code)) if isinstance(stock_code, float) else str(stock_code)
        code_str = code_str.zfill(6)
        
        print(f"正在获取股票 {code_str} 价格...")
        
        if code_str.startswith(('60', '68')):
            symbol = f"sh{code_str}"
        else:
            symbol = f"sz{code_str}"
        
        url = f"http://hq.sinajs.cn/list={symbol}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Referer": "http://finance.sina.com.cn/"
        }
        
        response = requests.get(url, headers=headers, timeout=10)
        response.encoding = 'gbk'
        
        pattern = r'="([^,]+),([^,]+),'
        match = re.search(pattern, response.text)
        
        if match:
            stock_name = match.group(1)
            price_value = float(match.group(2))
            print(f"✅ 股票 {code_str}({stock_name}) 价格: {price_value}")
            return price_value
        else:
            print(f"❌ 股票 {code_str} 未获取到价格数据")
    except Exception as e:
        print(f"❌ 获取股票 {stock_code} 价格失败: {e}")
    
    return None

def get_security_price(code) -> Optional[float]:
    """根据代码获取价格/净值"""
    if not code or pd.isna(code):
        return None
        
    code_str = str(int(code)) if isinstance(code, float) else str(code)
    code_str = code_str.zfill(6)
    
    if is_fund_code(code_str):
        return get_fund_nav(code_str)
    elif is_stock_code(code_str):
        return get_stock_price(code_str)
    else:
        print(f"⚠️  {code_str} 不是有效的基金或股票代码")
        return None

# ==================== 主逻辑函数 ====================

def update_excel_nav_values():
    """更新Excel中的净值列"""
    
    excel_path = "Asset.xlsx"
    
    if not os.path.exists(excel_path):
        print(f"❌ 文件 {excel_path} 不存在")
        return False
    
    try:
        # 读取Excel
        df = pd.read_excel(excel_path, sheet_name='data')
        print(f"📊 读取到 {len(df)} 行数据")
        
        updated_count = 0
        total_count = len(df)
        
        print("\n🔄 开始更新净值...")
        for idx, row in df.iterrows():
            code = row['基金代码']
            fund_name = row['基金名称']
            old_nav = row['净值']
            
            print(f"[{idx+1}/{total_count}] {fund_name}({code})")
            
            new_nav = get_security_price(code)  # 现在这个函数已经定义了
            
            if new_nav is not None and abs(new_nav - old_nav) > 0.0001:
                df.at[idx, '净值'] = new_nav
                print(f"   ✅ 更新: {old_nav:.4f} → {new_nav:.4f}")
                updated_count += 1
            elif new_nav is not None:
                print(f"   ℹ️  净值无变化: {old_nav:.4f}")
            else:
                print(f"   ❌ 获取失败，保持原值: {old_nav:.4f}")
            
            time.sleep(1)
        
        # 重新计算相关列
        df['持有金额'] = df['持有份额'] * df['净值']
        df['累计盈亏'] = df['持有金额'] - df['成本金额']
        df['持有收益率%'] = (df['累计盈亏'] / df['成本金额'] * 100).round(2)
        
        # 保存文件
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='data', index=False)
            
        print(f"\n🎉 更新完成! 更新了 {updated_count}/{total_count} 个净值")
        
        # 输出汇总信息
        total_value = df['持有金额'].sum()
        total_cost = df['成本金额'].sum()
        total_profit = total_value - total_cost
        profit_rate = (total_profit / total_cost * 100) if total_cost > 0 else 0
        
        print(f"\n📈 投资组合汇总:")
        print(f"总持有金额: {total_value:,.2f}元")
        print(f"总成本金额: {total_cost:,.2f}元")
        print(f"总累计盈亏: {total_profit:,.2f}元")
        print(f"总收益率: {profit_rate:.2f}%")
        
        return True
        
    except Exception as e:
        print(f"❌ 错误: {e}")
        import traceback
        traceback.print_exc()
        return False

# ==================== 主程序入口 ====================

if __name__ == "__main__":
    print("=" * 60)
    print("💰 投资组合净值自动更新工具")
    print(f"🕒 开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    
    success = update_excel_nav_values()
    
    if success:
        print("\n✨ 更新完成!")
    else:
        print("\n💥 更新失败")
    
    sys.exit(0 if success else 1)
