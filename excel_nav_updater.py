#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_nav_updater.py - 云端版本
"""

import pandas as pd
import requests
import re
import os
import sys
from typing import Optional
import time
from datetime import datetime

def is_github_actions():
    """判断是否在GitHub Actions环境中运行"""
    return os.getenv('GITHUB_ACTIONS') == 'true'

def update_excel_nav_values_cloud(excel_path: str = "Asset.xlsx"):
    """云端版本的净值更新"""
    
    print(f"🕒 开始执行净值更新: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
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
            
            new_nav = get_security_price(code)
            
            if new_nav is not None and abs(new_nav - old_nav) > 0.0001:
                df.at[idx, '净值'] = new_nav
                print(f"   ✅ 更新: {old_nav:.4f} → {new_nav:.4f}")
                updated_count += 1
            elif new_nav is not None:
                print(f"   ℹ️  净值无变化: {old_nav:.4f}")
            else:
                print(f"   ❌ 获取失败，保持原值: {old_nav:.4f}")
            
            time.sleep(1)  # 礼貌延迟
        
        # 重新计算
        df['持有金额'] = df['持有份额'] * df['净值']
        df['累计盈亏'] = df['持有金额'] - df['成本金额']
        df['持有收益率%'] = (df['累计盈亏'] / df['成本金额'] * 100).round(2)
        
        # 保存文件
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='data', index=False)
            
        print(f"\n🎉 更新完成! 更新了 {updated_count}/{total_count} 个净值")
        
        # 在GitHub Actions中输出总结
        if is_github_actions():
            total_value = df['持有金额'].sum()
            total_cost = df['成本金额'].sum()
            total_profit = total_value - total_cost
            profit_rate = (total_profit / total_cost * 100) if total_cost > 0 else 0
            
            print(f"::set-output name=total_value::{total_value:.2f}")
            print(f"::set-output name=total_profit::{total_profit:.2f}")
            print(f"::set-output name=profit_rate::{profit_rate:.2f}")
            
        return True
        
    except Exception as e:
        print(f"❌ 错误: {e}")
        return False

if __name__ == "__main__":
    success = update_excel_nav_values_cloud("Asset.xlsx")
    sys.exit(0 if success else 1)