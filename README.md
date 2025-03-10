# Markdown to DOCX 转换工具

这是一个由AI完全辅助开发的文档转换工具，用于将Markdown文件转换为带格式的Word文档。整个项目是通过提炼大模型交互提示词模式，然后直接基于提示词进行开发实现的。

## 项目起源

本项目源自对AI辅助开发的探索，通过以下步骤完成：

1. 首先定义了清晰的提示词模式（见 prompt_pattern.md）
2. 基于提示词模式与AI进行交互
3. AI根据提示词逐步实现整个项目

### 提示词模式核心要素

- 明确的需求定义
- 完整的项目结构设计
- 详细的实现计划
- 分步验证和优化
- 完整的文档说明

## 功能特点

- 支持Markdown到DOCX的转换
- 可自定义标题样式（字体、大小和粗细）
- 通过JSON配置文件灵活设置样式
- 完整的命令行参数支持

## 技术栈

- Python 3.7+
- python-docx：用于处理Word文档
- pypandoc：用于Markdown转换
- pandoc：核心转换引擎

## 安装步骤

1. 安装Python依赖：
```bash
pip install -r requirements.txt
```

2. 安装pandoc：
```bash
winget install pandoc
```

## 使用方法

基本使用：
```bash
python md2docx.py input.md output.docx
```

使用自定义样式：
```bash
python md2docx.py input.md output.docx -c style_config.json
```

## 配置文件说明

样式配置文件(style_config.json)示例：
```json
{
    "Heading 1": {
        "font": "宋体",
        "size": 16,
        "bold": true
    },
    "Heading 2": {
        "font": "宋体",
        "size": 14,
        "bold": true
    }
}
```

配置项说明：
- font: 字体名称
- size: 字号大小(磅值)
- bold: 是否加粗(true/false)

## AI辅助开发流程

本项目的开发完全基于AI辅助，遵循以下流程：

1. 提示词设计
   - 明确项目需求
   - 确认技术栈和依赖
   - 定义具体格式要求

2. 项目实现
   - AI根据提示词创建项目结构
   - 实现核心功能代码
   - 创建配置文件
   - 编写测试用例

3. 测试验证
   - 安装依赖
   - 运行测试
   - 验证功能
   - 优化改进

整个开发过程展示了如何通过精心设计的提示词模式，引导AI有效地完成软件开发任务。

## 项目结构

```
md2docx/
├── requirements.txt    # 项目依赖
├── md2docx.py         # 主程序
├── style_config.json  # 样式配置
├── test.md           # 测试文件
└── README.md         # 项目文档
```

## 注意事项

- 确保正确安装了pandoc
- 配置文件必须是有效的JSON格式
- 字体名称必须是系统中已安装的字体

## 项目特点

1. 完全AI辅助开发
2. 基于提示词模式构建
3. 代码清晰易懂
4. 配置灵活
5. 易于扩展

这个项目展示了如何通过系统化的提示词模式，引导AI完成从需求分析到代码实现的完整开发流程，为未来的AI辅助开发提供了一个实用的参考案例。