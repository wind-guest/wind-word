package markdown_test

import (
	"bytes"
	"io"
	"strings"
	"testing"

	"github.com/FengKeWG/wind-word/pkg/document"
	"github.com/FengKeWG/wind-word/pkg/markdown"
)

func TestExporterPreservesMathFromLiveDocument(t *testing.T) {
	doc := buildMathDocument()

	exporter := markdown.NewExporter(markdown.DefaultExportOptions())
	output, err := exporter.ExportToString(doc, nil)
	if err != nil {
		t.Fatalf("ExportToString failed: %v", err)
	}

	assertMarkdownMath(t, output)
}

func TestExporterPreservesMathFromOpenedDocument(t *testing.T) {
	doc := buildMathDocument()

	data, err := doc.ToBytes()
	if err != nil {
		t.Fatalf("ToBytes failed: %v", err)
	}

	opened, err := document.OpenFromMemory(io.NopCloser(bytes.NewReader(data)))
	if err != nil {
		t.Fatalf("OpenFromMemory failed: %v", err)
	}

	exporter := markdown.NewExporter(markdown.DefaultExportOptions())
	output, err := exporter.ExportToString(opened, nil)
	if err != nil {
		t.Fatalf("ExportToString failed: %v", err)
	}

	assertMarkdownMath(t, output)
}

func buildMathDocument() *document.Document {
	doc := document.New()

	para := doc.AddParagraph("Inline math: ")
	para.AddInlineMath(`a^2+b^2=c^2`)
	doc.AddMathFormula(`\frac{1}{n}\sum_{i=1}^{n} x_i`, true)

	return doc
}

func assertMarkdownMath(t *testing.T, output string) {
	t.Helper()

	normalized := strings.ReplaceAll(output, "\r\n", "\n")

	if !strings.Contains(normalized, "Inline math: $a^{2}+b^{2}=c^{2}$") {
		t.Fatalf("expected inline math in markdown, got:\n%s", normalized)
	}

	if !strings.Contains(normalized, "$$") ||
		!strings.Contains(normalized, `\frac{1}{n}\sum_{i=1}^{n}`) ||
		!strings.Contains(normalized, `x_{i}`) {
		t.Fatalf("expected block math in markdown, got:\n%s", normalized)
	}
}
