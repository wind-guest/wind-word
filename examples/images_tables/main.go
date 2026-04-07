package main

import (
	"fmt"
	"log"
	"os"

	"github.com/wind-guest/wind-word/examples/internal/exampleutil"
	"github.com/wind-guest/wind-word/pkg/document"
)

func main() {
	outputPath, err := exampleutil.OutputPath("images_tables.docx")
	if err != nil {
		log.Fatal(err)
	}
	sampleImagePath, err := exampleutil.OutputPath("sample.png")
	if err != nil {
		log.Fatal(err)
	}

	imageData, err := exampleutil.SamplePNGData(640, 320)
	if err != nil {
		log.Fatal(err)
	}
	if err := os.WriteFile(sampleImagePath, imageData, 0o644); err != nil {
		log.Fatal(err)
	}

	doc := document.New()
	doc.AddHeadingParagraph("图片与表格示例", 1)
	doc.AddParagraph("这个示例覆盖文件图片、内存图片、浮动图片、表格样式、合并单元格和单元格图片。")

	floatImage, err := doc.AddImageFromFile(sampleImagePath, &document.ImageConfig{
		Position: document.ImagePositionFloatRight,
		WrapText: document.ImageWrapSquare,
		Size: &document.ImageSize{
			Width:           55,
			KeepAspectRatio: true,
		},
		AltText: "Generated sample image",
		Title:   "Floating image",
		OffsetX: 3,
		OffsetY: 3,
	})
	if err != nil {
		log.Fatal(err)
	}
	if err := doc.SetImageTitle(floatImage, "Floating example"); err != nil {
		log.Fatal(err)
	}

	_, err = doc.AddImageFromData(imageData, "inline.png", document.ImageFormatPNG, 640, 320, &document.ImageConfig{
		Position:  document.ImagePositionInline,
		Alignment: document.AlignCenter,
		Size: &document.ImageSize{
			Width:           90,
			KeepAspectRatio: true,
		},
		AltText: "Inline generated image",
		Title:   "Inline image",
	})
	if err != nil {
		log.Fatal(err)
	}

	table, err := doc.AddTable(&document.TableConfig{
		Rows:  4,
		Cols:  3,
		Width: 9600,
	})
	if err != nil {
		log.Fatal(err)
	}

	_ = table.SetCellText(0, 0, "资源")
	_ = table.SetCellText(0, 1, "说明")
	_ = table.SetCellText(0, 2, "状态")
	_ = table.SetCellText(1, 0, "浮动图片")
	_ = table.SetCellText(1, 1, "来自文件路径")
	_ = table.SetCellText(1, 2, "ok")
	_ = table.SetCellText(2, 0, "单元格图片")
	_ = table.SetCellText(2, 1, "来自内存数据")
	_ = table.SetCellText(2, 2, "")
	_ = table.SetCellText(3, 0, "合并单元格")
	_ = table.SetCellText(3, 1, "下方会执行横向合并")
	_ = table.SetCellText(3, 2, "ready")

	if err := table.SetRowAsHeader(0, true); err != nil {
		log.Fatal(err)
	}
	if err := table.SetAlternatingRowColors("EAF2FF", "FFFFFF"); err != nil {
		log.Fatal(err)
	}
	if err := table.SetCellShading(0, 0, &document.ShadingConfig{
		Pattern:         document.ShadingPatternClear,
		ForegroundColor: "auto",
		BackgroundColor: "D9E8FB",
	}); err != nil {
		log.Fatal(err)
	}

	if _, err := doc.AddCellImageFromData(table, 2, 2, imageData, 30); err != nil {
		log.Fatal(err)
	}

	if err := table.MergeCellsHorizontal(3, 1, 2); err != nil {
		log.Fatal(err)
	}
	_ = table.SetCellText(3, 1, "这一行的后两列已经合并")

	if err := doc.Save(outputPath); err != nil {
		log.Fatal(err)
	}

	fmt.Println("wrote", outputPath)
}
