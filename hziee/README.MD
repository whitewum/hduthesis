# 杭州电子科技大学信息工程学院本科论文 LaTeX 模板

本项目基于 [`hduthesis`](https://github.com/myhsia/hduthesis) LaTeX 模板，适用于杭州电子科技大学信息工程学院本科毕业论文的撰写，严格遵循学校论文格式规范，支持中英文排版、数学公式、图表、参考文献等学术写作需求。

## 功能特点

- **合规性**：完全符合杭电信工学院本科论文格式要求（详见`格式要求.txt`）。
- **现代化**：采用 LaTeX3 语法，结构清晰，易于维护和扩展。
- **中英文混排**：自动处理中文和英文的字体、字号、行距等细节。
- **参考文献管理**：内置 BibTeX 支持，示例见 `hziee-bachelor.bib`。
- **高度可定制**：通过参数灵活配置标题、作者、专业、导师等信息。



- 其他：配置、说明、致谢等文件

## 快速开始
 
 可以在 Overleaf 复制一个新项目,[开始体验](https://cn.overleaf.com/read/sjfqfntmyctq#429731)


1. **准备环境**  
   推荐使用 TeX Live 2022 及以上版本，需支持 XeLaTeX 编译（以保证中文字体正确）。

2. **编译论文**  
   在 `hziee` 目录下，执行如下命令（可用 TeXworks、VSCode+LaTeX Workshop 或命令行）：

   ```bash
   xelatex hziee-bachelor.tex
   bibtex hziee-bachelor
   xelatex hziee-bachelor.tex
   xelatex hziee-bachelor.tex
   ```

   > 若需生成PDF并正确显示参考文献，需多次编译。

3. **自定义信息**  
   修改 `hziee-bachelor.tex` 文件开头的 `\hduset{...}` 参数，填写你的论文题目、作者、学号、专业、导师等信息。

4. **参考文献**  
   参考文献条目存放于 `hziee-bachelor.bib`，可按 BibTeX 规范增删。

5. **格式规范**  
   详细格式要求见 `格式要求.txt`，模板已自动实现绝大部分规范，无需手动调整。

## 其他用法

- **幻灯片/Beamer**  
  使用 `beamerthemehdu.sty` 可制作符合杭电风格的学术汇报幻灯片。
- **信纸/试卷**  
  支持生成学校信纸、试卷等文档，详见各模块代码和注释。

## 依赖说明

- XeLaTeX（推荐）
- 必要宏包：`ctex`, `biblatex`, `tikz`, `hyperref`, `amssymb`, `mathtools` 等
- 推荐完整安装 TeX Live 或 MacTeX

## 致谢

本模板基于 [myhsia/hduthesis](https://github.com/myhsia/hduthesis) 项目开发，并根据实际需求做了适配和修改。  
如有问题或建议，欢迎在原项目或本项目 issue 区反馈。

## Copyright and License

Copyright (C) 2023-2025 

It may be distributed and/or modified under the conditions of the LaTeX Project Public License (LPPL), either version 1.3c of this license or (at your option) any later version. The latest version of this license is in the file http://www.latex-project.org/lppl.txt

This work has the LPPL maintenance status `maintained`.

The Current Maintainer of this work is Min Wu.