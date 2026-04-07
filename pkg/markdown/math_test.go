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
		"# 数学公式示例",
		"",
		"## 欧拉公式",
		"",
		"$$e^{i\\pi} + 1 = 0$$",
		"",
		"欧拉公式是数学中最优美的公式之一，它将五个最重要的数学常数（e、i、π、1、0）联系在一起。",
		"",
		"## 二次方程求根公式",
		"",
		"$$x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}$$",
		"",
		"这是求解一元二次方程 $ax^2 + bx + c = 0$ 的标准公式。",
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
	if got := strings.Count(xml, "<m:oMathPara>"); got != 2 {
		t.Fatalf("expected 2 block math paragraphs, got %d:\n%s", got, xml)
	}
	if !strings.Contains(xml, "欧拉公式是数学中最优美的公式之一") {
		t.Fatalf("expected following Chinese paragraph to stay outside math block:\n%s", xml)
	}
	if !strings.Contains(xml, "二次方程求根公式") {
		t.Fatalf("expected later heading to remain in document.xml:\n%s", xml)
	}
	if strings.Contains(xml, "æ¬§") || strings.Contains(xml, "äºæ¬¡") {
		t.Fatalf("expected no mojibake in document.xml:\n%s", xml)
	}

	opened, err := document.OpenFromMemory(io.NopCloser(bytes.NewReader(data)))
	if err != nil {
		t.Fatalf("OpenFromMemory failed: %v", err)
	}

	output, err := markdown.NewExporter(markdown.DefaultExportOptions()).ExportToString(opened, nil)
	if err != nil {
		t.Fatalf("ExportToString failed: %v", err)
	}
	if !strings.Contains(output, "## **二次方程求根公式**") {
		t.Fatalf("expected second heading in exported markdown, got:\n%s", output)
	}
}

func TestMarkdownSingleLineDisplayMathInsideCodeFenceIsPreserved(t *testing.T) {
	input := strings.Join([]string{
		"```md",
		"$$e^{i\\pi} + 1 = 0$$",
		"```",
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
	if strings.Contains(xml, "<m:oMathPara>") {
		t.Fatalf("expected code fence content to stay as code text, got math paragraph:\n%s", xml)
	}
	if !strings.Contains(xml, "$$e^{i\\pi} + 1 = 0$$") {
		t.Fatalf("expected literal code fence text in document.xml, got:\n%s", xml)
	}
}

func TestMarkdownAdvancedLatexCommandsDoNotLeakRawCommandsIntoOMML(t *testing.T) {
	input := strings.Join([]string{
		"# Advanced",
		"",
		"$$i\\hbar \\frac{\\partial}{\\partial t}\\Psi(\\mathbf{r},t)=\\hat{H}\\Psi(\\mathbf{r},t)$$",
		"",
		"$$\\nabla \\cdot \\mathbf{E}=\\frac{\\rho}{\\varepsilon_0}, \\quad \\frac{d}{dt}\\left(\\frac{\\partial L}{\\partial \\dot{q}}\\right)=0$$",
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
	for _, raw := range []string{`\mathbf`, `\hat`, `\hbar`, `\varepsilon`, `\dot`} {
		if strings.Contains(xml, raw) {
			t.Fatalf("expected advanced command %s to be converted in document.xml, got:\n%s", raw, xml)
		}
	}
	for _, converted := range []string{"ℏ", "ϵ", "∇"} {
		if !strings.Contains(xml, converted) {
			t.Fatalf("expected converted symbol %s in document.xml, got:\n%s", converted, xml)
		}
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
