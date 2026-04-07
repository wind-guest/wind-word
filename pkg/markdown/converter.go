package markdown

import (
	"os"
	"path/filepath"
	"strings"

	"github.com/yuin/goldmark"
	"github.com/yuin/goldmark/extension"
	"github.com/yuin/goldmark/parser"
	"github.com/yuin/goldmark/text"

	mathjax "github.com/litao91/goldmark-mathjax"

	"github.com/wind-guest/wind-word/pkg/document"
)

// MarkdownConverter Markdown转换器接口
type MarkdownConverter interface {
	// ConvertFile 转换单个文件
	ConvertFile(mdPath, docxPath string, options *ConvertOptions) error

	// ConvertBytes 转换字节数据
	ConvertBytes(mdContent []byte, options *ConvertOptions) (*document.Document, error)

	// ConvertString 转换字符串
	ConvertString(mdContent string, options *ConvertOptions) (*document.Document, error)

	// BatchConvert 批量转换
	BatchConvert(inputs []string, outputDir string, options *ConvertOptions) error
}

// Converter 默认转换器实现
type Converter struct {
	md   goldmark.Markdown
	opts *ConvertOptions
}

// NewConverter 创建新的转换器实例
func NewConverter(opts *ConvertOptions) *Converter {
	if opts == nil {
		opts = DefaultOptions()
	}

	return &Converter{
		md:   buildMarkdown(opts),
		opts: opts.clone(),
	}
}

func buildMarkdown(opts *ConvertOptions) goldmark.Markdown {
	extensions := []goldmark.Extender{}
	if opts.EnableGFM {
		extensions = append(extensions, extension.GFM)
	}
	if opts.EnableFootnotes {
		extensions = append(extensions, extension.Footnote)
	}
	if opts.EnableMath {
		// 使用标准的LaTeX数学公式分隔符: $...$ 用于行内公式, $$...$$ 用于块级公式
		extensions = append(extensions, mathjax.NewMathJax(
			mathjax.WithInlineDelim("$", "$"),
			mathjax.WithBlockDelim("$$", "$$"),
		))
	}

	md := goldmark.New(
		goldmark.WithExtensions(extensions...),
		goldmark.WithParserOptions(
			parser.WithAutoHeadingID(),
		),
	)

	return md
}

// ConvertString 转换字符串内容为Word文档
func (c *Converter) ConvertString(content string, opts *ConvertOptions) (*document.Document, error) {
	return c.ConvertBytes([]byte(content), opts)
}

// ConvertBytes 转换字节数据为Word文档
func (c *Converter) ConvertBytes(content []byte, opts *ConvertOptions) (*document.Document, error) {
	resolved := c.opts.clone()
	if opts != nil {
		resolved = opts.clone()
	}

	if resolved.EnableMath {
		content = normalizeMarkdownMathBlocks(content)
	}

	// 创建新的Word文档
	doc := document.New()

	// 应用页面设置
	if resolved.PageSettings != nil {
		// 这里可以后续扩展，使用现有的页面设置API
	}

	// 解析Markdown
	reader := text.NewReader(content)
	md := buildMarkdown(resolved)
	astDoc := md.Parser().Parse(reader)

	// 创建渲染器并转换
	renderer := &WordRenderer{
		doc:    doc,
		opts:   resolved,
		source: content,
	}

	err := renderer.Render(astDoc)
	if err != nil {
		return nil, err
	}

	return doc, nil
}

// ConvertFile 转换文件
func (c *Converter) ConvertFile(mdPath, docxPath string, options *ConvertOptions) error {
	// 读取Markdown文件
	content, err := os.ReadFile(mdPath)
	if err != nil {
		return NewConversionError("FileRead", "failed to read markdown file", 0, 0, err)
	}

	// 设置图片基础路径（如果未指定）
	resolved := c.opts.clone()
	if options != nil {
		resolved = options.clone()
	}
	if resolved.ImageBasePath == "" {
		resolved.ImageBasePath = filepath.Dir(mdPath)
	}

	// 转换内容
	doc, err := c.ConvertBytes(content, resolved)
	if err != nil {
		return err
	}

	// 保存Word文档
	err = doc.Save(docxPath)
	if err != nil {
		return NewConversionError("FileSave", "failed to save word document", 0, 0, err)
	}

	return nil
}

func normalizeMarkdownMathBlocks(content []byte) []byte {
	if len(strings.TrimSpace(string(content))) == 0 {
		return content
	}

	newline := "\n"
	if strings.Contains(string(content), "\r\n") {
		newline = "\r\n"
	}

	text := strings.ReplaceAll(string(content), "\r\n", "\n")
	lines := strings.Split(text, "\n")
	normalized := make([]string, 0, len(lines))
	changed := false
	inFence := false
	fenceMarker := ""

	for _, line := range lines {
		trimmed := strings.TrimSpace(line)

		if marker, ok := markdownFenceMarker(trimmed); ok {
			if !inFence {
				inFence = true
				fenceMarker = marker
			} else if marker == fenceMarker {
				inFence = false
				fenceMarker = ""
			}
			normalized = append(normalized, line)
			continue
		}

		if inFence {
			normalized = append(normalized, line)
			continue
		}

		if replacement, ok := normalizeSingleLineMathBlock(line, trimmed); ok {
			normalized = append(normalized, replacement...)
			changed = true
			continue
		}

		normalized = append(normalized, line)
	}

	if !changed {
		return content
	}

	return []byte(strings.Join(normalized, newline))
}

func normalizeSingleLineMathBlock(rawLine, trimmed string) ([]string, bool) {
	if !strings.HasPrefix(trimmed, "$$") || !strings.HasSuffix(trimmed, "$$") || len(trimmed) <= 4 {
		return nil, false
	}

	inner := strings.TrimSpace(trimmed[2 : len(trimmed)-2])
	if inner == "" {
		return nil, false
	}

	indentLen := len(rawLine) - len(strings.TrimLeft(rawLine, " \t"))
	indent := rawLine[:indentLen]

	return []string{indent + "$$", indent + inner, indent + "$$"}, true
}

func markdownFenceMarker(trimmed string) (string, bool) {
	switch {
	case strings.HasPrefix(trimmed, "```"):
		return "```", true
	case strings.HasPrefix(trimmed, "~~~"):
		return "~~~", true
	default:
		return "", false
	}
}

// BatchConvert 批量转换文件
func (c *Converter) BatchConvert(inputs []string, outputDir string, options *ConvertOptions) error {
	// 确保输出目录存在
	err := os.MkdirAll(outputDir, 0755)
	if err != nil {
		return NewConversionError("DirectoryCreate", "failed to create output directory", 0, 0, err)
	}

	total := len(inputs)
	for i, input := range inputs {
		// 报告进度
		if options != nil && options.ProgressCallback != nil {
			options.ProgressCallback(i+1, total)
		}

		// 生成输出文件名
		base := strings.TrimSuffix(filepath.Base(input), filepath.Ext(input))
		output := filepath.Join(outputDir, base+".docx")

		// 转换单个文件
		err := c.ConvertFile(input, output, options)
		if err != nil {
			if options != nil && options.ErrorCallback != nil {
				options.ErrorCallback(err)
			}
			if options == nil || !options.IgnoreErrors {
				return err
			}
		}
	}

	return nil
}
