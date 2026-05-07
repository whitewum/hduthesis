# thesis-checker

杭州电子科技大学信息工程学院本科毕业设计/论文 Word 格式检查 skills。

这个 skill 用于辅助检查 `.docx` 论文的格式规范性，并给出结构化中文报告。检查结果会区分：

- `critical`：严重问题，必须修改
- `warning`：一般问题，建议修改
- `passed`：已通过检查项
- `content_notes`：内容质量参考评估

当前版本为测试版，适合先小范围给同学或老师试用反馈。

## 功能

- 页眉内容、字号、对齐检查
- 摘要、Abstract、目录、参考文献结构检查
- 一级、二级、三级标题格式检查
- 正文字号、行距、首行缩进检查
- 图题、表题、图表引用格式检查
- 参考文献与正文引用格式检查
- 内容质量参考扫描，例如字数、章节结构、文献引用密度

## 工作方式

本项目采用 XML 解包检查流程：

1. 使用 `docx/scripts/office/unpack.py` 解包 `.docx`
2. 使用 `thesis-checker/scripts/check_format_xml.py` 解析 Word OOXML
3. 输出 JSON 检查结果
4. 由 AI 按 `SKILL.md` 整理为中文报告

检查脚本只使用 Python 标准库，不依赖 `python-docx`。

## 安装

### 方式一：导入打包好的 skill

下载或使用仓库中的：

```text
thesis-checker.skill
```

在支持 skills 的客户端中导入该文件。

### 方式二：手动安装

将 `thesis-checker/` 目录复制到客户端的 skills 目录中，目录结构应类似：

```text
thesis-checker/
├── SKILL.md
├── scripts/
│   └── check_format_xml.py
└── references/
    └── format_rules.md
```

同时需要确保环境中有 `docx/scripts/office/unpack.py`。本仓库已附带 `docx/` 目录供本地测试使用。

## 使用

上传或提供一篇 `.docx` 论文后，对 AI 说：

```text
请使用 thesis-checker 检查这篇论文格式，并列出 critical 和 warning。
```

报告应包含：

- 严重问题完整列表
- 一般问题完整列表
- 已通过检查项摘要
- 内容质量参考评估
- 总体修改建议

## 本地测试

在仓库根目录运行：

```bash
python3 docx/scripts/office/unpack.py "论文.docx" /tmp/thesis_unpacked/
python3 thesis-checker/scripts/check_format_xml.py /tmp/thesis_unpacked/
```

脚本会输出 JSON，例如：

```json
{
  "critical": [],
  "warning": [],
  "passed": [],
  "content_notes": []
}
```

## 重新打包

修改 `thesis-checker/` 后，运行：

```bash
zip -r /tmp/thesis-checker.skill thesis-checker/ \
  -x "*/__pycache__/*" "*.pyc" "*.DS_Store"
cp /tmp/thesis-checker.skill ./thesis-checker.skill
```

## 仓库结构

```text
.
├── README.md
├── AGENTS.md
├── 格式要求.txt
├── thesis-checker/
│   ├── SKILL.md
│   ├── scripts/
│   │   └── check_format_xml.py
│   └── references/
│       └── format_rules.md
├── thesis-checker.skill
└── docx/
    └── scripts/office/unpack.py
```

## 隐私说明

不要把真实论文、学生姓名、学号、导师信息等隐私材料提交到公开仓库。

本仓库的 `.gitignore` 默认忽略：

- `*.docx`
- `reports/`
- `_skills_tmp/`
- Python 缓存和系统临时文件

## 已知问题

- 只在ClaudeCode和Codex测试，未在其他AI AGENT测试过。
- 只在mac进行过测试，未在windows测试过
- 当前为规则检查工具，不能替代人工最终审核。
- 内容质量评估只是参考，不应作为专业水平的绝对判断。
- 多页眉或多节文档中，个别 `passed` 汇总项仍可能需要进一步收紧。

