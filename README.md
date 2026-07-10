# 标准格式文件输出 Skill

这是一个用于生成标准 Word 文件的 Codex Skill。它内置两套经过保留和分析的 `.docx` 模板，可根据文件用途选择输出模式。

## 支持的模式

### 公文模式（official）

适用于单位内部正式文件，例如：

- 通知、请示、报告、函、纪要；
- 制度、办法和正式工作材料；
- 需要使用单位内部标准公文版式的 Word 文件。

公文模式使用 `assets/gongwen.docx`，保留原模板的页面设置、字体、字号、标题层级、自动编号、缩进和行距。

### 通用模式（general）

适用于对外协作和一般交流材料，例如：

- 合作方案、交流材料、沟通说明；
- 项目汇报、参考资料、情况介绍；
- 不属于正式公文，但需要统一专业版式的 Word 文件。

通用模式使用 `assets/reference.docx`，保留其标题、副标题、正文及多级标题样式。

## 安装

将仓库克隆或下载到 Codex Skills 目录：

```text
~/.codex/skills/standard-document-output/
```

安装后，可在提示词中显式调用：

```text
使用 $standard-document-output，以公文模式生成一份工作通知。
```

```text
使用 $standard-document-output，把这份合作交流材料按通用模式输出为 Word。
```

也可以不指定模式，让 Skill 根据用途判断：

```text
帮我把这份内容生成标准格式 Word，请根据用途选择合适模板。
```

## Markdown 输入约定

生成脚本接受 UTF-8 Markdown：

```markdown
# 文档主标题
> 通用模式副标题
## 一级标题
正文段落。
### 二级标题
正文段落。
#### 三级标题
##### 四级标题
```

公文模式的层级编号由 Word 模板自动生成，不要在 Markdown 中重复输入“一、”“（一）”“1.”或“（1）”。

## 命令行生成

安装 `python-docx` 后运行：

```bash
python scripts/build_standard_doc.py --mode official input.md output.docx
```

或：

```bash
python scripts/build_standard_doc.py --mode general input.md output.docx
```

## 目录结构

```text
standard-document-output/
├── SKILL.md
├── README.md
├── agents/
│   └── openai.yaml
├── assets/
│   ├── gongwen.docx
│   └── reference.docx
├── references/
│   ├── official-format.md
│   └── general-format.md
└── scripts/
    ├── build_standard_doc.py
    └── build_gongwen.py
```

## 使用原则

- 两个 Word 模板是最终格式基准，不使用近似样式替代。
- 不直接修改模板资产；每次生成时复制模板并写入内容。
- 用户明确指定模式时服从用户；没有指定时根据文件用途选择。
- 正式交付前应将生成的 Word 渲染为 PDF 或页面图片，逐页检查乱码、分页、编号、表格和版面问题。

## 隐私提示

模板和生成文件可能包含作者、单位、联系人或其他元数据。公开分享生成文件前，请根据实际情况检查并清理敏感信息。
