package document_test

import (
	"archive/zip"
	"bytes"
	"io"
	"strings"
	"testing"

	"github.com/FengKeWG/wind-word/pkg/document"
)

func TestMathAPIsProduceOMML(t *testing.T) {
	doc := document.New()

	para := doc.AddParagraph("Inline math: ")
	para.AddInlineMath(`a^2+b^2=c^2`)
	doc.AddMathFormula(`\frac{1}{n}\sum_{i=1}^{n} x_i`, true)

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
	if strings.Contains(xml, "[Formula]") {
		t.Fatalf("placeholder formula text leaked into document.xml:\n%s", xml)
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
