# Resume Word Tailor

简历匹配与 Word 母版编辑 skill：根据岗位 JD 提炼真实经历，并用 Microsoft Word 在已有简历模板中替换指定正文，输出可编辑的 `.docx`。

适合已有满意排版、需要反复定向投递的场景。支持旧版 `.doc` 和 `.docx` 作为输入；不附带任何个人简历、照片或预设工作经历。

## 能做什么

- 将 JD 关键词映射到已有经历，保留实习和项目结构。
- 将已定稿文字原样写入指定段落。
- 保留段落样式、标题与正文加粗、页边距和浮动图片几何信息。
- 保存后重新打开，验证替换内容、未修改的段落、分页和格式快照。
- 拒绝覆盖已有输出；源文件以只读方式打开，并核对源文件 SHA-256。
- 模板文字不匹配、分页超限、表格段落或复杂正文格式时停止，避免擅自缩字、删句或改变布局。

不负责制造经历、预测简历过筛概率，也不将个人项目成果表述成企业上线收益。

## 环境

Windows、PowerShell 7、已安装且可正常启动的桌面版 Microsoft Word。普通执行不需要 Python、Office 在线服务或额外 npm 包。

脚本通过 Word COM 执行。Linux、macOS、仅安装 WPS 或只有网页版 Word 的环境不支持这个执行方式。执行期间会启动独立的隐藏 Word 实例；请避免同时编辑相同文件。

## 安装到 Codex

克隆后将 `skills/resume-word-tailor` 复制到个人 skills 目录。以下命令会在目标已存在时停止，防止覆盖本地个性化配置：

```powershell
git clone https://github.com/elvis239/resume-word-tailor.git
$sourceSkill = Join-Path (Get-Location) 'resume-word-tailor\skills\resume-word-tailor'
$targetSkill = Join-Path $env:USERPROFILE '.codex\skills\resume-word-tailor'
if (Test-Path -LiteralPath $targetSkill) { throw 'Skill already exists; review before updating.' }
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $targetSkill) | Out-Null
Copy-Item -LiteralPath $sourceSkill -Destination $targetSkill -Recurse
```

显式调用：

> 使用 $resume-word-tailor，根据这个 JD 调整我的简历，按我指定的 Word 母版输出。

> 使用 $resume-word-tailor，把上面确定的文字原样写入 Word，保留格式。

自动选择已启用。触发范围是“简历匹配并输出 Word”或“定稿文字写入简历模板”；普通面试问答和文本建议不触发。

## 手动执行

在 `skills/resume-word-tailor` 目录下运行。示例使用合成内容，真实 inspection 和 change plan 包含简历正文，应保留在本机。

先读取模板：

```powershell
& .\scripts\edit-resume.ps1 -Mode Inspect `
  -TemplatePath 'C:\resumes\template.doc' `
  -InspectionPath 'C:\workspace\.tmp\inspection.json'
```

每个段落都有当前序号、文字与格式。根据检查结果创建 UTF-8 JSON：

```json
{
  "replacements": [
    {
      "index": 3,
      "expectedText": "技术栈：SQL",
      "text": "技术栈：SQL、Python"
    }
  ]
}
```

再写入新的输出：

```powershell
& .\scripts\edit-resume.ps1 -Mode Apply `
  -TemplatePath 'C:\resumes\template.doc' `
  -PlanPath 'C:\workspace\.tmp\changes.json' `
  -OutputPath 'C:\workspace\output\resume-role.docx'
```

可选 `-PreviewPdf 'C:\workspace\.tmp\preview.pdf'` 用原生 Word 导出布局检查文件。默认页数上限等于模板实际页数；仅在用户同意扩页后使用 `-MaxPages`。

`expectedText` 必须与当前模板逐字一致，且不含末尾段落标记。`text` 不可包含换段或单元格标记。修改模板后重新 Inspect，不沿用旧的段落序号。

## 验证与局限

执行原生 Word 集成测试：

```powershell
& .\tests\smoke-word.ps1
```

测试使用临时合成模板，检查正常替换、模板不匹配拒绝、分页超限拒绝和原文件不变；测试材料放在仓库 `.tmp/` 并清理。

脚本适合普通段落型模板。带正文局部多样化加粗、表格、域代码、超链接或特殊控件的模板需要进一步定制，不保证泛化处理。图片验证包括数量与几何信息，不是每个嵌入对象的逐字节比对。原生 Word 首次视觉检查仍然必要；自动验证不保证与其他 Office/WPS 版本分页完全一致。

个人模板路径和经历事实应存放在本机 `AGENTS.md` 或其他本地资料中，不能随 skill 打包发布。仓库忽略文档、图片、生成结果、暂存目录和私有配置。
