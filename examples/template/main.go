package main

import (
	"fmt"
	"log"

	"github.com/FengKeWG/wind-word/examples/internal/exampleutil"
	"github.com/FengKeWG/wind-word/pkg/document"
)

func main() {
	templatePath, err := exampleutil.OutputPath("template_source.docx")
	if err != nil {
		log.Fatal(err)
	}
	outputPath, err := exampleutil.OutputPath("template_rendered.docx")
	if err != nil {
		log.Fatal(err)
	}

	if err := buildTemplate(templatePath); err != nil {
		log.Fatal(err)
	}

	renderer := document.NewTemplateRenderer()
	renderer.SetLogging(false)

	if _, err := renderer.LoadTemplateFromFile("weekly_report", templatePath); err != nil {
		log.Fatal(err)
	}

	logoData, err := exampleutil.SamplePNGData(360, 180)
	if err != nil {
		log.Fatal(err)
	}

	data := document.NewTemplateData()
	data.SetVariable("header_title", "wind-word weekly report")
	data.SetVariable("owner", "FengKeWG")
	data.SetCondition("enabled", true)
	data.SetList("items", []interface{}{
		map[string]interface{}{"name": "清理旧代码", "status": "done"},
		map[string]interface{}{"name": "整理示例", "status": "in progress"},
		map[string]interface{}{"name": "修复冒烟问题", "status": "done"},
	})
	data.SetImageFromData("logo", logoData, &document.ImageConfig{
		Position:  document.ImagePositionInline,
		Alignment: document.AlignCenter,
		Size: &document.ImageSize{
			Width:           50,
			KeepAspectRatio: true,
		},
		AltText: "Generated logo",
		Title:   "Generated logo",
	})

	doc, err := renderer.RenderTemplate("weekly_report", data)
	if err != nil {
		log.Fatal(err)
	}

	if err := doc.Save(outputPath); err != nil {
		log.Fatal(err)
	}

	fmt.Println("wrote", templatePath)
	fmt.Println("wrote", outputPath)
}

func buildTemplate(path string) error {
	doc := document.New()
	if err := doc.AddHeader(document.HeaderFooterTypeDefault, "{{header_title}}"); err != nil {
		return err
	}

	title := doc.AddFormattedParagraph("项目周报模板", &document.TextFormat{
		Bold:      true,
		FontSize:  18,
		FontColor: "1F4B99",
	})
	title.SetAlignment(document.AlignCenter)

	doc.AddParagraph("负责人：{{owner}}")
	doc.AddParagraph("模板状态：{{#if enabled}}已启用{{else}}未启用{{/if}}")
	doc.AddParagraph("任务列表：")
	doc.AddParagraph("{{#each items}}")
	doc.AddParagraph("- {{name}} / {{status}}")
	doc.AddParagraph("{{/each}}")
	doc.AddParagraph("标识图片：")
	doc.AddParagraph("{{#image logo}}")

	return doc.Save(path)
}
