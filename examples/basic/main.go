package main

import (
	"fmt"
	"log"

	"github.com/wind-guest/wind-word/examples/internal/exampleutil"
	"github.com/wind-guest/wind-word/pkg/document"
	stylepkg "github.com/wind-guest/wind-word/pkg/style"
)

func main() {
	outputPath, err := exampleutil.OutputPath("basic.docx")
	if err != nil {
		log.Fatal(err)
	}

	doc := document.New()
	if err := doc.SetTitle("wind-word Basic Example"); err != nil {
		log.Fatal(err)
	}
	if err := doc.SetAuthor("FengKeWG"); err != nil {
		log.Fatal(err)
	}
	if err := doc.SetDescription("Basic example for wind-word"); err != nil {
		log.Fatal(err)
	}
	if err := doc.SetPageOrientation(document.OrientationPortrait); err != nil {
		log.Fatal(err)
	}
	if err := doc.SetPageMargins(25, 20, 25, 20); err != nil {
		log.Fatal(err)
	}
	if err := doc.AddHeader(document.HeaderFooterTypeDefault, "wind-word basic example"); err != nil {
		log.Fatal(err)
	}
	if err := doc.AddFooterWithPageNumber(document.HeaderFooterTypeDefault, "Page ", true); err != nil {
		log.Fatal(err)
	}

	styleAPI := stylepkg.NewQuickStyleAPI(doc.GetStyleManager())
	_, err = styleAPI.CreateQuickStyle(stylepkg.QuickStyleConfig{
		ID:      "Callout",
		Name:    "Callout",
		Type:    stylepkg.StyleTypeParagraph,
		BasedOn: "Normal",
		ParagraphConfig: &stylepkg.QuickParagraphConfig{
			LineSpacing: 1.2,
			SpaceBefore: 8,
			SpaceAfter:  8,
		},
		RunConfig: &stylepkg.QuickRunConfig{
			FontName:  "Calibri",
			FontSize:  11,
			FontColor: "24557A",
			Bold:      true,
		},
	})
	if err != nil {
		log.Fatal(err)
	}

	title := doc.AddFormattedParagraph("wind-word", &document.TextFormat{
		Bold:       true,
		FontSize:   20,
		FontColor:  "1F4B99",
		FontFamily: "Calibri",
	})
	title.SetAlignment(document.AlignCenter)

	doc.AddParagraph("这个示例展示基础文档创建、页面设置、标题、列表、脚注、目录、样式和表格能力。")

	callout := doc.AddParagraph("这个段落应用了自定义样式 Callout。")
	callout.SetStyle("Callout")

	doc.AddHeadingWithBookmark("第一章 概述", 1, "chapter_1")
	para := doc.AddParagraph("wind-word 当前专注于 .docx 核心读写、结构化编辑和持续维护。")
	para.SetSpacing(&document.SpacingConfig{
		LineSpacing:     1.5,
		BeforePara:      6,
		AfterPara:       6,
		FirstLineIndent: 24,
	})

	if err := doc.AddFootnote("这是一段带脚注的正文", "脚注内容：这里用于演示脚注能力。"); err != nil {
		log.Fatal(err)
	}

	doc.AddHeadingWithBookmark("第二章 列表", 1, "chapter_2")
	doc.AddBulletList("支持无序列表", 0, document.BulletTypeDot)
	doc.AddBulletList("支持多级列表", 1, document.BulletTypeCircle)
	doc.AddNumberedList("支持有序列表", 0, document.ListTypeDecimal)

	doc.AddPageBreak()

	doc.AddHeadingWithBookmark("第三章 表格", 1, "chapter_3")
	table, err := doc.AddTable(&document.TableConfig{
		Rows:  4,
		Cols:  3,
		Width: 9000,
	})
	if err != nil {
		log.Fatal(err)
	}

	_ = table.SetCellText(0, 0, "字段")
	_ = table.SetCellText(0, 1, "值")
	_ = table.SetCellText(0, 2, "状态")
	_ = table.SetCellText(1, 0, "Repository")
	_ = table.SetCellText(1, 1, "wind-word")
	_ = table.SetCellText(1, 2, "active")
	_ = table.SetCellText(2, 0, "Branch")
	_ = table.SetCellText(2, 1, "main")
	_ = table.SetCellText(2, 2, "tracked")
	_ = table.SetCellText(3, 0, "备注")
	_ = table.SetCellText(3, 1, "示例输出已生成")
	_ = table.SetCellText(3, 2, "ok")

	if err := table.SetRowAsHeader(0, true); err != nil {
		log.Fatal(err)
	}
	if err := table.ApplyTableStyle(&document.TableStyleConfig{
		FirstRowHeader:    true,
		FirstColumnHeader: true,
		BandedRows:        true,
	}); err != nil {
		log.Fatal(err)
	}

	tocConfig := document.DefaultTOCConfig()
	tocConfig.Title = "目录"
	tocConfig.MaxLevel = 2
	if err := doc.GenerateTOC(tocConfig); err != nil {
		log.Fatal(err)
	}

	if err := doc.Save(outputPath); err != nil {
		log.Fatal(err)
	}

	fmt.Println("wrote", outputPath)
}
