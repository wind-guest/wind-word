package markdown_test

import (
	"archive/zip"
	"bytes"
	"io"
	"strings"
	"testing"

	"github.com/wind-guest/wind-word/pkg/document"
	"github.com/wind-guest/wind-word/pkg/markdown"
)

func TestMarkdownMathProducesOMML(t *testing.T) {
	input := strings.Join([]string{
		"# Math",
		"",
		"Inline: $a^2+b^2=c^2$",
		"",
		"$$",
		`\frac{1}{n}\sum_{i=1}^{n} x_i`,
		"$$",
	}, "\n")

	converter := markdown.NewConverter(markdown.DefaultOptions())
	doc, err := converter.ConvertString(input, nil)
	if err != nil {
		t.Fatalf("ConvertString failed: %v", err)
	}

	data, err := doc.ToBytes()
	if err != nil {
		t.Fatalf("ToBytes failed: %v", err)
	}

	xml := readZipPart(t, data, "word/document.xml")
	if !strings.Contains(xml, "<m:oMath>") {
		t.Fatalf("expected inline OMML in document.xml, got:\n%s", xml)
	}
	if !strings.Contains(xml, "<m:oMathPara>") {
		t.Fatalf("expected block OMML in document.xml, got:\n%s", xml)
	}
	if strings.Contains(xml, "Cambria Math") {
		t.Fatalf("math fell back to styled text instead of OMML:\n%s", xml)
	}
}

func TestMarkdownSingleLineDisplayMathDoesNotCaptureFollowingSections(t *testing.T) {
	input := strings.Join([]string{
		"# 数学公式示例文档",
		"",
		"## 第一部分：基本公式",
		"",
		"### 行间公式",
		"",
		"求和公式：",
		"",
		"$$\\int_a^b f(x)dx$$",
		"",
		"---",
		"",
		"## 第二部分：极限公式",
		"",
		"### 极限表达",
		"",
		"$$\\lim_{x\\to\\infty} \\frac{1}{x} = 0$$",
		"",
		"---",
		"",
		"## 第三部分：矩阵示例",
		"",
		"### 矩阵公式",
		"",
		"$$\\begin{bmatrix} a & b \\\\ c & d \\end{bmatrix}$$",
	}, "\n")

	converter := markdown.NewConverter(markdown.DefaultOptions())
	doc, err := converter.ConvertString(input, nil)
	if err != nil {
		t.Fatalf("ConvertString failed: %v", err)
	}

	data, err := doc.ToBytes()
	if err != nil {
		t.Fatalf("ToBytes failed: %v", err)
	}

	xml := readZipPart(t, data, "word/document.xml")
	if got := strings.Count(xml, "<m:oMathPara>"); got != 3 {
		t.Fatalf("expected 3 block math paragraphs, got %d:\n%s", got, xml)
	}
	if !strings.Contains(xml, "第二部分：极限公式") {
		t.Fatalf("expected later heading to stay outside math block:\n%s", xml)
	}
	if strings.Contains(xml, "ç¬¬") {
		t.Fatalf("expected no garbled UTF-8 sequences in document.xml:\n%s", xml)
	}
	if !strings.Contains(xml, "<m:m>") {
		t.Fatalf("expected matrix OMML in document.xml:\n%s", xml)
	}
	if strings.Contains(xml, `\begin`) {
		t.Fatalf("matrix fell back to raw LaTeX text in document.xml:\n%s", xml)
	}

	opened, err := document.OpenFromMemory(io.NopCloser(bytes.NewReader(data)))
	if err != nil {
		t.Fatalf("OpenFromMemory failed: %v", err)
	}

	output, err := markdown.NewExporter(markdown.DefaultExportOptions()).ExportToString(opened, nil)
	if err != nil {
		t.Fatalf("ExportToString failed: %v", err)
	}

	if !strings.Contains(output, "## **第二部分：极限公式**") {
		t.Fatalf("expected second section heading in exported markdown, got:\n%s", output)
	}
	if !strings.Contains(output, `\begin{bmatrix}`) {
		t.Fatalf("expected matrix to survive markdown export, got:\n%s", output)
	}
}

func readZipPart(t *testing.T, data []byte, part string) string {
	t.Helper()

	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("zip.NewReader failed: %v", err)
	}

	for _, file := range reader.File {
		if file.Name != part {
			continue
		}
		rc, err := file.Open()
		if err != nil {
			t.Fatalf("open part %s failed: %v", part, err)
		}
		defer rc.Close()

		content, err := io.ReadAll(rc)
		if err != nil {
			t.Fatalf("read part %s failed: %v", part, err)
		}
		return string(content)
	}

	t.Fatalf("part %s not found", part)
	return ""
}
