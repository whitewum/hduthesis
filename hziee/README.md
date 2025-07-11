# 杭州电子科技大学信息工程学院本科毕业设计LaTeX模板

本模板基于 [hduthesis](https://github.com/myhsia/hduthesis) 修改，专门适配杭州电子科技大学信息工程学院本科毕业设计的格式要求。

## 特性

- 完全符合信息工程学院本科毕业设计格式要求
- 支持中英文混排
- 自动生成目录、参考文献
- 支持数学公式、图表、代码等
- 兼容Overleaf在线编译

## 使用方法

### 在线使用（推荐）

1. 点击 Overleaf模板链接 直接使用
2. 复制模板到您的Overleaf账户
3. 修改 `hziee-bachelor.tex` 中的个人信息
4. 开始撰写您的毕业设计

### 本地使用

1. 下载所有文件到本地
2. 确保已安装完整的LaTeX发行版（推荐TeX Live）
3. 使用XeLaTeX编译 `hziee-bachelor.tex`

```bash
xelatex hziee-bachelor.tex
bibtex hziee-bachelor
xelatex hziee-bachelor.tex
xelatex hziee-bachelor.tex
```

## 文件说明

- `hduthesis.cls` - 主模板文件
- `hziee-bachelor.tex` - 示例文档主文件
- `hziee-bachelor.bib` - 参考文献数据库
- `hdu-*.pdf` - 学校标识图片文件
- `beamerthemehdu.sty` - 演示文稿主题（可选）

## 个人信息设置

在 `hziee-bachelor.tex` 文件开头的 `\hduset` 命令中修改以下信息：

```latex
\hduset
  {
    title      = 您的毕业设计题目,
    department = 您的学院,
    major      = 您的专业,
    class      = 您的班级,
    stdntid    = 您的学号,
    author     = 您的姓名,
    supervisor = 您的指导教师,
    bibsource  = hziee-bachelor
  }
```

## 格式要求

本模板严格按照信息工程学院的格式要求制作：

- 封面：28号宋体加粗标题，楷体小三号个人信息
- 页眉：五号宋体"杭州电子科技大学信息工程学院本科毕业设计"
- 一级标题：黑体三号居中
- 二级标题：黑体四号左对齐
- 三级标题：黑体小四号左对齐
- 正文：中文宋体、西文Times New Roman，小四号，固定行距20磅
- 图表：五号字体，行距固定值20磅

## 许可证

本项目基于 LaTeX Project Public License (LPPL) 1.3c 发布。

## 致谢

- 感谢 [hduthesis](https://github.com/myhsia/hduthesis) 项目提供的基础模板
- 感谢杭州电子科技大学信息工程学院提供的格式规范

## 问题反馈

如果您在使用过程中遇到问题，请通过以下方式反馈：

- 提交 GitHub Issue


## 更新日志

### v1.0.0 (2025-07-11)

- 初始版本发布
- 适配信息工程学院格式要求
- 支持Overleaf在线编译
