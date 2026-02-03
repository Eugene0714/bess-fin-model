#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel版本同步脚本
@description 监听网页版文件变化，自动更新Excel版本
@version 1.0
"""

import os
import time
from datetime import datetime
from pathlib import Path
import hashlib

def get_file_hash(filepath):
    """计算文件MD5哈希值"""
    if not os.path.exists(filepath):
        return None
    with open(filepath, 'rb') as f:
        return hashlib.md5(f.read()).hexdigest()

def check_files_changed():
    """检查关键文件是否发生变化"""
    key_files = [
        'financial-model.js',
        'index.html',
    ]
    
    hash_file = '.excel_sync_hash.json'
    current_hashes = {}
    
    # 读取之前的哈希值
    old_hashes = {}
    if os.path.exists(hash_file):
        import json
        with open(hash_file, 'r', encoding='utf-8') as f:
            old_hashes = json.load(f)
    
    # 计算当前哈希值
    changed = False
    for file in key_files:
        if os.path.exists(file):
            current_hash = get_file_hash(file)
            current_hashes[file] = current_hash
            if file not in old_hashes or old_hashes[file] != current_hash:
                changed = True
                print(f"检测到文件变化: {file}")
    
    # 保存当前哈希值
    if changed:
        import json
        with open(hash_file, 'w', encoding='utf-8') as f:
            json.dump(current_hashes, f, indent=2)
    
    return changed

def sync_excel():
    """同步生成Excel文件"""
    if check_files_changed():
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 检测到文件变化，开始生成Excel...")
        try:
            from generate_excel import create_excel_file
            create_excel_file()
            print("✓ Excel文件已同步更新")
        except Exception as e:
            print(f"✗ 生成Excel文件时出错: {e}")
    else:
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 文件未变化，无需更新")

if __name__ == "__main__":
    sync_excel()
