# wind-word

`wind-word` 是一个用于创建、读取、修改 `.docx` 文件的 Go 库，模块地址为 `github.com/wind-guest/wind-word`。

## 安装

要求：Go `1.22+`

```bash
go get github.com/wind-guest/wind-word
```

常用包：

```go
import "github.com/wind-guest/wind-word/pkg/document"
import "github.com/wind-guest/wind-word/pkg/markdown"
import "github.com/wind-guest/wind-word/pkg/style"
```

## 主要能力

- 创建、打开、保存 `.docx`
- 段落、标题、分页符、文本格式设置
- 页面尺寸、页边距、页方向、页眉页脚距离设置
- 页眉、页脚、页码
- 表格创建、单元格格式、合并、嵌套表格
- 图片插入、浮动图片、单元格图片
- 目录、书签、标题跳转
- 脚注、尾注、列表编号
- 模板渲染，支持变量、条件、循环、图片占位符
- Markdown -> Word
- Word -> Markdown
- 样式管理与自定义样式

## 仓库结构

- `pkg/document`: Word 文档核心读写与编辑能力
- `pkg/markdown`: Markdown 与 Word 双向转换
- `pkg/style`: 样式系统与样式工具

当前仓库只保留核心库代码，历史示例、验证脚本和外围目录已经移除，README 作为主要入口说明。

## 快速开始

### 创建一个文档

```go
package main

import (
	"log"

	"github.com/wind-guest/wind-word/pkg/document"
)

func main() {
	doc := document.New()

	title := doc.AddFormattedParagraph("wind-word", &document.TextFormat{
		Bold:      true,
		FontSize:  18,
		FontColor: "1F4B99",
	})
	title.SetAlignment(document.AlignCenter)

	doc.AddParagraph("这是一个由 Go 生成的 Word 文档。")
	doc.AddHeadingParagraph("功能示例", 1)

	table, err := doc.AddTable(&document.TableConfig{
		Rows:  2,
		Cols:  2,
		Width: 9000,
	})
	if err != nil {
		log.Fatal(err)
	}

	_ = table.SetCellText(0, 0, "字段")
	_ = table.SetCellText(0, 1, "值")
	_ = table.SetCellText(1, 0, "Status")
	_ = table.SetCellText(1, 1, "Active")

	if err := doc.Save("example.docx"); err != nil {
		log.Fatal(err)
	}
}
```

### 页面、页眉页脚、目录

```go
doc := document.New()

_ = doc.SetPageOrientation(document.OrientationLandscape)
_ = doc.SetPageMargins(25, 20, 25, 20)

_ = doc.AddHeader(document.HeaderFooterTypeDefault, "wind-word")
_ = doc.AddFooterWithPageNumber(document.HeaderFooterTypeDefault, "Page ", true)

doc.AddHeadingWithBookmark("第一章 概述", 1, "chapter_1")
doc.AddHeadingWithBookmark("1.1 背景", 2, "chapter_1_1")

tocConfig := document.DefaultTOCConfig()
tocConfig.Title = "目录"
tocConfig.MaxLevel = 2
_ = doc.GenerateTOC(tocConfig)
```

### 图片、脚注、列表

```go
doc := document.New()

_, _ = doc.AddImageFromFile("cover.png", &document.ImageConfig{
	Position:  document.ImagePositionInline,
	Alignment: document.AlignCenter,
})

_ = doc.AddFootnote("这是正文内容", "这是脚注内容")

doc.AddBulletList("第一项", 0, document.BulletTypeDot)
doc.AddNumberedList("第二项", 0, document.ListTypeDecimal)
```

## Markdown 转换

### Markdown 转 Word

```go
converter := markdown.NewConverter(markdown.DefaultOptions())
err := converter.ConvertFile("input.md", "output.docx", nil)
```

### Word 转 Markdown

```go
exporter := markdown.NewExporter(markdown.DefaultExportOptions())
err := exporter.ExportToFile("input.docx", "output.md", nil)
```

## 模板能力

模板引擎支持这些常见占位形式：

- `{{name}}`
- `{{#if enabled}} ... {{/if}}`
- `{{#each items}} ... {{/each}}`
- `{{#image logo}}`

典型流程：

```go
renderer := document.NewTemplateRenderer()

_, _ = renderer.LoadTemplateFromFile("report", "template.docx")

data := document.NewTemplateData()
data.SetVariable("name", "wind-word")
data.SetCondition("enabled", true)
data.SetList("items", []interface{}{
	map[string]interface{}{"title": "Item A"},
	map[string]interface{}{"title": "Item B"},
})
data.SetImage("logo", "logo.png", nil)

doc, err := renderer.RenderTemplate("report", data)
if err != nil {
	log.Fatal(err)
}

_ = doc.Save("report.docx")
```

## 日志

库的默认日志输出是静默的，不会主动向标准输出打印内部日志。

如果你需要排查问题，可以手动开启：

```go
document.SetGlobalLevel(document.LogLevelDebug)
```

## 致敬原作者

在此致敬原作者 [zerx-lab/wordZero](https://github.com/zerx-lab/wordZero)。

`wind-word` 基于他的工作继续演进，但已经在此基础上进行了大幅修改和优化，后续也会持续更新和维护这个库。

## License

见 `LICENSE`。
