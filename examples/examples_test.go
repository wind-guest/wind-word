package examples_test

import (
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"testing"

	"github.com/FengKeWG/wind-word/examples/internal/exampleutil"
)

func TestExamplesRun(t *testing.T) {
	_, file, _, ok := runtime.Caller(0)
	if !ok {
		t.Fatal("failed to resolve examples directory")
	}

	examplesDir := filepath.Dir(file)
	cases := []struct {
		name    string
		outputs []string
	}{
		{
			name:    "basic",
			outputs: []string{"basic.docx"},
		},
		{
			name:    "images_tables",
			outputs: []string{"sample.png", "images_tables.docx"},
		},
		{
			name:    "template",
			outputs: []string{"template_source.docx", "template_rendered.docx"},
		},
		{
			name:    "markdown",
			outputs: []string{"markdown_input.md", "markdown.docx", "markdown_roundtrip.md"},
		},
		{
			name:    "math",
			outputs: []string{"math.docx"},
		},
	}

	for _, tc := range cases {
		tc := tc
		t.Run(tc.name, func(t *testing.T) {
			for _, output := range tc.outputs {
				outputPath, err := exampleutil.OutputPath(output)
				if err != nil {
					t.Fatalf("resolve output path %s: %v", output, err)
				}
				_ = os.Remove(outputPath)
			}

			cmd := exec.Command("go", "run", "./"+tc.name)
			cmd.Dir = examplesDir
			output, err := cmd.CombinedOutput()
			if err != nil {
				t.Fatalf("go run %s failed: %v\n%s", tc.name, err, output)
			}
			if !strings.Contains(string(output), "wrote ") {
				t.Fatalf("unexpected example output for %s: %s", tc.name, output)
			}

			for _, expected := range tc.outputs {
				outputPath, err := exampleutil.OutputPath(expected)
				if err != nil {
					t.Fatalf("resolve output path %s: %v", expected, err)
				}
				if _, err := os.Stat(outputPath); err != nil {
					t.Fatalf("expected output file %s was not created: %v", outputPath, err)
				}
			}
		})
	}
}
