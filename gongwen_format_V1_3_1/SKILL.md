---
name: gongwen_format_V1_3_1
description: Use this skill whenever the user wants to format Chinese official documents, public-sector materials, reports, notices, summaries, or pasted text into a standardized .docx; asks for 公文排版, 材料排版, 按公文格式, 生成 Word/docx, 清除原格式, 标题空一行, 正文两端对齐, 中英文标点适配, 重点字段强调, or wants default official-document typography with optional temporary overrides. This skill should also be used when the user needs local font matching for 方正小标宋-style title fonts or wants to verify a generated .docx layout.
---

# 公文排版助手

Use this skill to turn plain text, `.txt`, or `.docx` into a standardized Chinese official-document `.docx`. The default behavior should match the user's desktop tool unless the user explicitly overrides it.

## Core Workflow

1. Determine the input source: pasted text, `.txt`, or `.docx`.
2. Start from `assets/default_template.json`.
3. Apply any temporary user overrides from the current request.
4. Check local fonts and resolve the title font keyword `方正小标宋`.
5. Generate the `.docx` with `scripts/format_gongwen.py`.
6. Report the output path, actual title-font match, missing fonts, and any important overrides used.
7. When accuracy matters, verify the result with `scripts/inspect_docx.py`.

Do not bundle font files. The skill only writes font names into the Word document and checks whether matching fonts appear to exist on the user's computer.

## Default Formatting

Use these defaults when the user does not specify otherwise:

- Page: A4.
- Margins: top 36 mm, bottom 36 mm, left 27 mm, right 27 mm.
- Line spacing: fixed 30 pt.
- Document title: title font keyword `方正小标宋`, size `二号`, centered.
- Level-1 heading: `黑体`, size `小二`, first-line indent 2 characters.
- Level-2 heading: `楷体_GB2312`, size `小二`, first-line indent 2 characters.
- Body: `仿宋_GB2312`, size `小二`, first-line indent 2 characters, justified.
- After the document title, insert one blank paragraph with body formatting and no text.
- Default focus fields: `分析认为`, `一是`, `二是`, `三是`, `四是`, `五是`, `六是`, `七是`, `八是`, `九是`, `十是`.
- Default focus-field style: `仿宋_GB2312`, size `小二`, bold.
- Normalize punctuation by language context unless the user asks not to.

## Font Matching

For the document title, treat `方正小标宋` as a font intent rather than a fixed exact font name. Search installed local font names for candidates containing that keyword, such as:

- `方正小标宋`
- `方正小标宋_GBK`
- `方正小标宋 GBK`
- `方正小标宋简体`

Use the closest match automatically, but tell the user what was selected. If no match is found, warn the user and still generate the document with the requested font name unless they choose another font.

For other default fonts (`黑体`, `楷体_GB2312`, `仿宋_GB2312`), check whether the exact name appears to exist locally. If a font is missing, warn the user that Word may substitute it and invite them to specify a replacement.

## User Overrides

Users may temporarily override formatting in natural language. Preserve unspecified defaults.

Common examples:

- `标题字体改成宋体`
- `正文改成仿宋，小三`
- `页边距改成上35下35左28右26`
- `行距改成28磅`
- `重点字段改成：重点关注、特别说明、经研究`
- `在默认重点字段基础上追加：风险提示、需要说明`
- `重点字段用黑体、小二、不加粗`
- `不要自动转换标点`
- `导出到 D:\材料输出`

If the user says "重点字段改成", replace the default focus-field list. If the user says "追加", append to the default list while removing duplicates.

## Scripts

Generate from a file:

```powershell
python .\scripts\format_gongwen.py --input "D:\input.docx" --output-dir "D:\out"
```

Generate from a text file:

```powershell
python .\scripts\format_gongwen.py --text-file "D:\input.txt" --output-dir "D:\out"
```

Use temporary overrides:

```powershell
python .\scripts\format_gongwen.py --input "D:\input.docx" --output-dir "D:\out" --line-spacing-pt 28 --body-font "宋体" --body-size "小三" --append-focus-fields "风险提示|需要说明"
```

Use a custom config:

```powershell
python .\scripts\format_gongwen.py --input "D:\input.docx" --config "D:\custom_template.json" --output-dir "D:\out"
```

Inspect a generated document:

```powershell
python .\scripts\inspect_docx.py "D:\out\材料_排版结果.docx"
```

## Dependencies

The formatter requires Python and `python-docx`.

If `python-docx` is missing, ask the user whether to install it, then run:

```powershell
python -m pip install python-docx
```

## Reference

Read `references/formatting_rules.md` when changing the rules, explaining behavior, or debugging document structure, punctuation normalization, focus-field matching, or font resolution.
