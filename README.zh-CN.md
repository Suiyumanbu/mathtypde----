# 面向 Codex 的 MathType Word skill

[English](README.md)

`mathtype-word` 可以把 Codex 生成的 LaTeX 公式插入 Word 文档的指定位置，并保留为可编辑的 MathType 公式。它也能直接为已有 MathType 显示公式添加原生右编号，而不重建 OLE 内容。

项目遵循 [OpenAI skill 文件夹格式](https://learn.chatgpt.com/docs/build-skills)：可安装内容位于 `mathtype-word/`，入口是 `SKILL.md`，脚本和参考文档与其一起分发。

## 功能

- 通过 MathType 自带的 TeX 转换器把 LaTeX 转为可编辑的 `Equation.DSMT4` OLE 对象。
- 使用唯一文本锚点或 Word 书签精确定位。
- 支持 `inline`、`display` 和 `right-numbered` 三种模式。
- 可按书签或 MathType 公式序号把已有显示公式转为原生右编号，且不重新执行 TeX 转换。
- 右编号使用 MathType 原生 `MTEqn`/`MTPlaceRef` 字段，兼容 MathType 后续编号管理命令。
- 以只读方式打开源文档，不覆盖已有输出；全部成功后才发布结果。
- 生成结构化 JSON 报告，并可导出 PDF 供逐页检查。
- 不使用模拟键盘输入或剪贴板。

## 环境要求

- Windows 桌面会话
- Microsoft Word 桌面版
- 已安装 Word 命令的 MathType 7
- 已启用本地 skills 的 Codex

当前版本已在 64 位 Word 2024 与 MathType 7.11 上完成端到端验证。其他较新的 Word/MathType 7 组合预计也可工作，建议先运行 `-Probe` 和仓库中的集成测试。

## 安装

### 使用 Release 压缩包

从仓库的 **Releases** 页面下载 ZIP，解压后运行：

```powershell
& .\install.ps1
```

若设置了 `CODEX_HOME`，安装器会使用 `$CODEX_HOME/skills`；否则依次选择已经存在的 `~/.agents/skills`、`~/.codex/skills`，都不存在时使用 `~/.agents/skills`。可用 `-DestinationRoot` 指定其他目录。默认不会覆盖已有安装；使用 `-Upgrade` 时，旧版本会先保存为带时间戳的备份。

### 手动安装

把完整的 `mathtype-word` 文件夹复制到 Codex skills 目录，并保留以下结构：

```text
skills/
└── mathtype-word/
    ├── SKILL.md
    ├── VERSION
    ├── LICENSE
    ├── THIRD_PARTY_NOTICES.md
    ├── agents/
    ├── references/
    └── scripts/
```

如果 Codex 没有立即发现它，请重启 Codex 或新建一个任务。

## 在 Codex 中使用

先在生成的 DOCX 中放置唯一锚点或书签，再让 Codex 调用此 skill；已有 MathType 公式也可以按序号定位。例如：

```text
使用 $mathtype-word，把 [[MT:energy]] 替换为 E=mc^2 的内联 MathType
公式，把 [[MT:quadratic]] 替换为带 MathType 原生右编号的求根公式；
同时导出 PDF 并检查结果。
```

Codex 会创建 JSON 清单并运行随附的 PowerShell 桥接脚本。也可以直接执行：

```powershell
& .\mathtype-word\scripts\insert-equations.ps1 `
  -Manifest .\examples\equations.example.json
```

清单中的路径以 JSON 文件所在目录为基准。请先把 `examples/equations.example.json` 复制到输入 DOCX 旁边，或修改其中的路径。

| 模式 | 结果 | 定位要求 |
|---|---|---|
| `inline` | 位于正文中的公式，或单独成段的内联 OLE 对象 | 锚点或书签 |
| `display` | 居中的 MathType 显示公式 | 定位内容必须独占段落 |
| `right-numbered` | 显示公式，并在右侧加入 MathType 管理的编号 | 定位内容必须独占段落 |

完整格式见[清单说明](mathtype-word/references/manifest.md)和 [LaTeX 兼容性说明](mathtype-word/references/latex-compatibility.md)。

给已有显示公式补编号时，可以使用包围其 OLE 对象的书签，或它在全部 MathType 公式中的一基序号：

```json
{
  "input": "draft-with-display-equations.docx",
  "output": "draft-with-native-numbers.docx",
  "equations": [
    {"bookmark": "mt_display_energy", "operation": "number-existing"},
    {"equationIndex": 4, "operation": "number-existing"}
  ]
}
```

正式运行本身会完成清单和本机环境校验；`-ValidateOnly` 与 `-Probe` 仅用于首次配置、CI 或故障诊断。快速迭代时可以省略 `pdf`，需要版面验收时再导出。

## 校验与测试

静态校验不需要 Word 或 MathType：

```powershell
python tools/validate_release.py
& .\scripts\validate.ps1
```

集成测试会先插入七个公式，再在第二次运行中为两个已有显示公式补编号。测试会确认 OLE 内容未被重建、后续编号顺序正确、无关字段保持不变，并成功导出两份 PDF。该测试需要本机安装 Word 和 MathType：

```powershell
python -m pip install -r tests/requirements.txt
& .\tests\run-integration.ps1
```

## 构建发行包

```powershell
& .\scripts\package-release.ps1
```

命令会生成 `dist/mathtype-word-<version>.zip` 及对应的 SHA-256 文件。给 GitHub 提交添加 `v<version>` 标签后，随附的工作流会自动创建 GitHub Release 并上传这两个文件。

下载后可在 PowerShell 中运行 `Get-FileHash .\mathtype-word-<version>.zip -Algorithm SHA256`，并与相邻 `.sha256` 文件中的值核对。

## 许可证与产品名称

本项目使用 [MIT License](LICENSE)。Microsoft Word 与 MathType 是用户自行安装的第三方产品，本项目不分发这些产品；详见[第三方声明](THIRD_PARTY_NOTICES.md)。
