// Package document 提供Word文档的核心操作功能
package document

import (
	"encoding/xml"

	"github.com/FengKeWG/wind-word/pkg/omml"
)

type OfficeMath = omml.OfficeMath
type OfficeMathPara = omml.OfficeMathPara

// MathParagraph 表示包含数学公式的段落
// 用于在文档中嵌入数学公式
type MathParagraph struct {
	XMLName    xml.Name             `xml:"w:p"`
	Properties *ParagraphProperties `xml:"w:pPr,omitempty"`
	Math       *OfficeMath          `xml:"m:oMath,omitempty"`
	MathPara   *OfficeMathPara      `xml:"m:oMathPara,omitempty"`
	Runs       []Run                `xml:"w:r"`
}

// MarshalXML 自定义序列化
func (mp *MathParagraph) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "w:p"}
	// 开始段落元素
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// 序列化段落属性
	if mp.Properties != nil {
		if err := e.Encode(mp.Properties); err != nil {
			return err
		}
	}

	// 序列化Runs（在公式之前的文本）
	for _, run := range mp.Runs {
		if err := e.Encode(run); err != nil {
			return err
		}
	}

	// 序列化数学公式（块级）
	if mp.MathPara != nil {
		if err := e.Encode(mp.MathPara); err != nil {
			return err
		}
	}

	// 序列化数学公式（行内）
	if mp.Math != nil {
		if err := e.Encode(mp.Math); err != nil {
			return err
		}
	}

	// 结束段落元素
	return e.EncodeToken(start.End())
}

// ElementType 返回数学段落元素类型
func (mp *MathParagraph) ElementType() string {
	return "math_paragraph"
}

// AddMathFormula 向文档添加数学公式
// latex: LaTeX格式的数学公式
// isBlock: 是否为块级公式（true为块级，false为行内）
func (d *Document) AddMathFormula(latex string, isBlock bool) *MathParagraph {
	Debugf("添加数学公式: %s (块级: %v)", latex, isBlock)

	mp := &MathParagraph{
		Runs: []Run{},
	}

	math := omml.LaTeXToOMML(latex)
	if isBlock {
		mp.MathPara = &OfficeMathPara{
			Math: math,
		}
	} else {
		mp.Math = math
	}

	d.Body.Elements = append(d.Body.Elements, mp)
	return mp
}

// AddInlineMath 在段落末尾添加一个行内公式。
func (p *Paragraph) AddInlineMath(latex string) {
	Debugf("向段落添加行内数学公式")

	if latex == "" {
		return
	}

	run := Run{
		Math: omml.LaTeXToOMML(latex),
	}
	p.Runs = append(p.Runs, run)
}
