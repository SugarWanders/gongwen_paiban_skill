# 公文排版助手 Skill V1_3_1

这是公文排版助手配套的 skill 独立发布目录，可供支持本地 skill、规则文档或脚本调用的 AI 编程工具使用。

## 版本信息

- Skill 名称：`gongwen_format_V1_3_1`
- 对应应用执行文件：`公文排版助手V1_3_1.exe`
- 版本标识：`V1_3_1`

## 字体说明

本 skill 运行时仅依赖用户本地电脑中已经安装的字体，并在生成 `.docx` 时写入对应字体名称。

如果本地电脑缺少某个字体，Word 可能会使用系统替代字体显示。需要精确字体效果时，请先在本机安装对应字体。

## 目录内容

- `gongwen_format_V1_3_1/`：可安装的 skill 目录。
- `archives/gongwen_format_V1_3_1.zip`：压缩后的 skill 包。
- `README.md`：本说明。
- `功能说明.md`：skill 可实现功能说明。
- `版本说明.md`：本版本发布说明。

## 安装方式

将 `gongwen_format_V1_3_1/` 复制到所用工具的本地 skill 或规则目录。以 Codex 为例：

```powershell
Copy-Item -Recurse .\gongwen_format_V1_3_1 C:\Users\MIA\.codex\skills\
```

如果使用 Claude Code 等其他工具，请按对应工具的本地规则目录或项目上下文方式引用本目录；也可以直接调用 `gongwen_format_V1_3_1/scripts/format_gongwen.py` 生成 `.docx`。

本目录用于单独的 skill 仓库，不随 EXE 应用仓库一起提交。
