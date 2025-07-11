# CLAUDE.md

本文件为Claude Code（claude.ai/code）在本仓库工作时提供操作指引。

## 项目概述

本仓库包含两个高校的LaTeX论文模板：
- **HDU**（杭州电子科技大学）：完整的LaTeX类模板
- **HZIEE**（杭州电子科技大学信息工程学院）：带有特定格式要求的论文模板

## 架构说明

### HDU模板结构
- `hdu/hduthesis.dtx`：主LaTeX文档源文件，包含完整类定义
- `hdu/hduthesis.ins`：用于从.dtx生成类文件的安装脚本
- `hdu/hduthesis-bachelor.pdf`：生成的本科论文模板PDF
- `hdu/hdutitle.pdf`：封面模板PDF

### HZIEE模板结构
- `hziee/hziee-bachelor.tex`：主论文文档，使用`hzieethesis`文档类
- `hziee/格式要求.txt`：详细格式要求说明

## 关键技术细节

### HDU模板
- 基于`LaTeX-expl3`现代LaTeX编程实现
- 支持本科与硕士论文格式
- 模块化设计，主要模块包括：
  - `typeset`：数学与文本排版
  - `layout`：封面与浮动体布局
  - `bc.config`：本科论文配置
  - `pg.config`：硕士论文配置
  - `beamer`：HDU Beamer幻灯片主题
  - `stationery`：校用信纸生成
  - `exam`：考试解答排版

### HZIEE模板
- 使用`hzieethesis`文档类（需自行创建）
- 字体要求：宋体SC（自动加粗）、华文黑体
- 数学字体：STIX Two Math
- 文档结构包括：封面、摘要（中英文）、目录、章节、参考文献、附录

## 构建命令

### HDU模板
```bash
cd hdu/
# 从文档源生成类文件
pdflatex hduthesis.ins
# 构建文档
pdflatex hduthesis.dtx
pdflatex hduthesis.dtx  # 需运行两次以保证交叉引用正确
```

### HZIEE模板
```bash
cd hziee/
# 构建论文主文档
pdflatex hziee-bachelor.tex
bibtex hziee-bachelor    # 如使用参考文献
pdflatex hziee-bachelor.tex
pdflatex hziee-bachelor.tex  # 多次运行以保证交叉引用正确
```

## HZIEE格式要求

摘自`格式要求.txt`的主要排版规范：
- **封面**：宋体28号加粗，居中
- **填写内容**：楷体小三号
- **页眉**：宋体五号，封面后各页均居中
- **中文摘要**：黑体三号标题，宋体小四内容，20pt行距
- **英文摘要**：Times New Roman三号加粗标题，Times New Roman小四内容
- **章标题**：黑体三号，居中，上下各空两行
- **节标题**：黑体四号（二级），黑体小四号（三级），左对齐
- **正文**：宋体小四号，20pt固定行距，首行缩进2字符
- **图表**：宋体/Times New Roman五号，20pt行距
- **页码**：宋体五号，阿拉伯数字

## 开发说明

- HDU模板为完整、带注释的LaTeX类，可作为参考实现
- HZIEE模板当前使用的`hzieethesis`类尚未实现，需自行开发
- 两套模板均遵循中国学术论文排版规范，但具体要求不同
- 字体处理对中文排版尤为关键