package main

import (
	"fmt"
	"log"
	"os"

	"github.com/wind-guest/wind-word/examples/internal/exampleutil"
	"github.com/wind-guest/wind-word/pkg/markdown"
)

func main() {
	inputPath, err := exampleutil.OutputPath("markdown_input.md")
	if err != nil {
		log.Fatal(err)
	}
	docxPath, err := exampleutil.OutputPath("markdown.docx")
	if err != nil {
		log.Fatal(err)
	}
	roundtripPath, err := exampleutil.OutputPath("markdown_roundtrip.md")
	if err != nil {
		log.Fatal(err)
	}

	source := `# wind-word Markdown Example

这是一个 Markdown 到 Word 再导回 Markdown 的示例。

## Quote

> 这是一段引用文本。

## Task List

- [x] 整理仓库
- [ ] 扩展示例

## Table

| Feature | Status |
| --- | --- |
| DOCX | Ready |
| Template | Active |

## Code

` + "```go" + `
fmt.Println("wind-word")
` + "```" + `

## Math

Inline math: $a^2+b^2=c^2$

$$
\frac{1}{n} \sum_{i=1}^{n} x_i
$$
`

	if err := os.WriteFile(inputPath, []byte(source), 0o644); err != nil {
		log.Fatal(err)
	}

	converter := markdown.NewConverter(markdown.DefaultOptions())
	if err := converter.ConvertFile(inputPath, docxPath, nil); err != nil {
		log.Fatal(err)
	}

	exporter := markdown.NewExporter(markdown.DefaultExportOptions())
	if err := exporter.ExportToFile(docxPath, roundtripPath, nil); err != nil {
		log.Fatal(err)
	}

	fmt.Println("wrote", inputPath)
	fmt.Println("wrote", docxPath)
	fmt.Println("wrote", roundtripPath)
}
