#!/usr/bin/env python
# -*- coding: utf-8 -*-

import argparse
import json
import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import pypandoc

def load_style_config(config_path):
    """加载样式配置文件"""
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def apply_styles(doc, style_config):
    """应用文档样式"""
    # 设置标题样式
    for i in range(1, 10):  # 支持1-9级标题
        style_name = f'Heading {i}'
        if style_name in style_config:
            style = doc.styles[style_name]
            config = style_config[style_name]
            
            # 设置字体
            if 'font' in config:
                style.font.name = config['font']
            # 设置字号
            if 'size' in config:
                style.font.size = Pt(config['size'])
            # 设置粗体
            if 'bold' in config:
                style.font.bold = config['bold']

def convert_md_to_docx(input_file, output_file, style_config_path=None):
    """将Markdown文件转换为Word文档"""
    # 首先使用pandoc将md转换为docx
    pypandoc.convert_file(
        input_file,
        'docx',
        outputfile=output_file,
        format='md'
    )
    
    # 如果提供了样式配置，则应用自定义样式
    if style_config_path and os.path.exists(style_config_path):
        doc = Document(output_file)
        style_config = load_style_config(style_config_path)
        apply_styles(doc, style_config)
        doc.save(output_file)

def main():
    parser = argparse.ArgumentParser(description='将Markdown文件转换为Word文档')
    parser.add_argument('input', help='输入的Markdown文件路径')
    parser.add_argument('output', help='输出的Word文档路径')
    parser.add_argument('-c', '--config', help='样式配置文件路径')
    
    args = parser.parse_args()
    
    try:
        convert_md_to_docx(args.input, args.output, args.config)
        print(f'转换成功！文件已保存为: {args.output}')
    except Exception as e:
        print(f'转换失败: {str(e)}')

if __name__ == '__main__':
    main()