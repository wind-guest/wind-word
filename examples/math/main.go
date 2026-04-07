package main

import (
	"fmt"
	"log"

	"github.com/wind-guest/wind-word/examples/internal/exampleutil"
	"github.com/wind-guest/wind-word/pkg/document"
)

func main() {
	outputPath, err := exampleutil.OutputPath("math.docx")
	if err != nil {
		log.Fatal(err)
	}

	doc := document.New()
	if err := doc.SetTitle("wind-word Math Example"); err != nil {
		log.Fatal(err)
	}
	if err := doc.SetAuthor("FengKeWG"); err != nil {
		log.Fatal(err)
	}
	if err := doc.SetDescription("Complex formula example for wind-word"); err != nil {
		log.Fatal(err)
	}

	title := doc.AddFormattedParagraph("wind-word Formula Handbook", &document.TextFormat{
		Bold:       true,
		FontSize:   20,
		FontColor:  "1F4B99",
		FontFamily: "Calibri",
	})
	title.SetAlignment(document.AlignCenter)

	intro := doc.AddParagraph("这个示例集中展示 direct API 的公式能力，包含行内公式、块级公式、连续公式段落和表格中的公式。")
	intro.SetSpacing(&document.SpacingConfig{
		LineSpacing: 1.4,
		AfterPara:   10,
	})

	doc.AddHeadingParagraph("1. 行内公式", 1)

	p1 := doc.AddParagraph("勾股定理：")
	p1.AddInlineMath(`a^2+b^2=c^2`)
	p1.AddFormattedText("，这是最常见的直角三角形关系。", nil)

	p2 := doc.AddParagraph("欧拉恒等式：")
	p2.AddInlineMath(`e^{i\pi}+1=0`)
	p2.AddFormattedText("，以及变量序列 ", nil)
	p2.AddInlineMath(`x_1+x_2+x_3+x_4`)
	p2.AddFormattedText("。", nil)

	p3 := doc.AddParagraph("希腊字母示例：")
	p3.AddInlineMath(`\alpha+\beta=\gamma`)
	p3.AddFormattedText("，概率表达式：", nil)
	p3.AddInlineMath(`P(A|B)=\frac{P(A\cap B)}{P(B)}`)

	doc.AddHeadingParagraph("2. 块级公式", 1)

	doc.AddParagraph("一元二次方程求根公式：")
	centerMath(doc.AddMathFormula(`x=\frac{-b+\sqrt{b^2-4ac}}{2a}`, true))

	doc.AddParagraph("三维向量长度的立方根形式：")
	centerMath(doc.AddMathFormula(`\sqrt[3]{x_1^2+x_2^2+x_3^2}`, true))

	doc.AddParagraph("算术平均值：")
	centerMath(doc.AddMathFormula(`\frac{a_1+a_2+\cdots+a_n}{n}`, true))

	doc.AddParagraph("积分示例：")
	centerMath(doc.AddMathFormula(`\int_0^1 x^2 dx`, true))

	doc.AddHeadingParagraph("3. 连续公式段落", 1)

	chain := doc.AddParagraph("下面是一段连续公式组合：")
	chain.AddFormattedText("当 ", nil)
	chain.AddInlineMath(`f(x)=\frac{1}{1+x^2}`)
	chain.AddFormattedText(" 时，其导数可以写成 ", nil)
	chain.AddInlineMath(`f'(x)=\frac{-2x}{(1+x^2)}`)
	chain.AddFormattedText("，再结合 ", nil)
	chain.AddInlineMath(`x_0^2+x_1^2+x_2^2`)
	chain.AddFormattedText(" 做局部分析。", nil)

	doc.AddPageBreak()

	doc.AddHeadingParagraph("4. 表格中的公式", 1)

	table, err := doc.AddTable(&document.TableConfig{
		Rows:  4,
		Cols:  3,
		Width: 9600,
	})
	if err != nil {
		log.Fatal(err)
	}

	_ = table.SetCellText(0, 0, "类别")
	_ = table.SetCellText(0, 1, "公式")
	_ = table.SetCellText(0, 2, "说明")
	_ = table.SetCellText(1, 0, "代数")
	_ = table.SetCellText(2, 0, "概率")
	_ = table.SetCellText(3, 0, "数列")
	_ = table.SetCellText(1, 2, "二项式展开中的基础关系")
	_ = table.SetCellText(2, 2, "条件概率形式")
	_ = table.SetCellText(3, 2, "有限项求和")

	if err := table.SetRowAsHeader(0, true); err != nil {
		log.Fatal(err)
	}
	if err := table.SetAlternatingRowColors("F2F7FF", "FFFFFF"); err != nil {
		log.Fatal(err)
	}

	addFormulaCell(table, 1, 1, `(a+b)^2=a^2+2ab+b^2`)
	addFormulaCell(table, 2, 1, `P(A|B)=\frac{P(A\cap B)}{P(B)}`)
	addFormulaCell(table, 3, 1, `S_n=\frac{n(a_1+a_n)}{2}`)

	doc.AddHeadingParagraph("5. 说明", 1)
	doc.AddParagraph("当前示例优先覆盖已经接入 OMML 转换链路的常用语法：分数、根号、上下标、希腊字母、积分和常见运算符。")
	doc.AddParagraph("更复杂的 LaTeX 语法后续还会继续增强。")

	if err := doc.Save(outputPath); err != nil {
		log.Fatal(err)
	}

	fmt.Println("wrote", outputPath)
}

func centerMath(mp *document.MathParagraph) {
	mp.Properties = &document.ParagraphProperties{
		Justification: &document.Justification{Val: string(document.AlignCenter)},
	}
}

func addFormulaCell(table *document.Table, row, col int, latex string) {
	para, err := table.AddCellParagraph(row, col, "")
	if err != nil {
		log.Fatal(err)
	}
	para.AddInlineMath(latex)
}
