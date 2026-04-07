package omml

import (
	"encoding/xml"
	"strings"
)

var unicodeToLaTeXToken = map[string]string{
	"α": `\alpha`,
	"β": `\beta`,
	"γ": `\gamma`,
	"δ": `\delta`,
	"ε": `\epsilon`,
	"ϵ": `\varepsilon`,
	"ζ": `\zeta`,
	"η": `\eta`,
	"θ": `\theta`,
	"ι": `\iota`,
	"κ": `\kappa`,
	"λ": `\lambda`,
	"μ": `\mu`,
	"ν": `\nu`,
	"ξ": `\xi`,
	"π": `\pi`,
	"ρ": `\rho`,
	"σ": `\sigma`,
	"τ": `\tau`,
	"υ": `\upsilon`,
	"φ": `\phi`,
	"χ": `\chi`,
	"ψ": `\psi`,
	"ω": `\omega`,
	"Α": `\Alpha`,
	"Β": `\Beta`,
	"Γ": `\Gamma`,
	"Δ": `\Delta`,
	"Ε": `\Epsilon`,
	"Ζ": `\Zeta`,
	"Η": `\Eta`,
	"Θ": `\Theta`,
	"Ι": `\Iota`,
	"Κ": `\Kappa`,
	"Λ": `\Lambda`,
	"Μ": `\Mu`,
	"Ν": `\Nu`,
	"Ξ": `\Xi`,
	"Π": `\Pi`,
	"Ρ": `\Rho`,
	"Σ": `\Sigma`,
	"Τ": `\Tau`,
	"Υ": `\Upsilon`,
	"Φ": `\Phi`,
	"Χ": `\Chi`,
	"Ψ": `\Psi`,
	"Ω": `\Omega`,
	"×": `\times`,
	"÷": `\div`,
	"±": `\pm`,
	"∓": `\mp`,
	"·": `\cdot`,
	"∗": `\ast`,
	"⋆": `\star`,
	"∘": `\circ`,
	"∙": `\bullet`,
	"⊕": `\oplus`,
	"⊖": `\ominus`,
	"⊗": `\otimes`,
	"⊘": `\oslash`,
	"⊙": `\odot`,
	"≤": `\leq`,
	"≥": `\geq`,
	"≠": `\neq`,
	"≈": `\approx`,
	"≡": `\equiv`,
	"∼": `\sim`,
	"≃": `\simeq`,
	"≅": `\cong`,
	"∝": `\propto`,
	"≪": `\ll`,
	"≫": `\gg`,
	"⊂": `\subset`,
	"⊃": `\supset`,
	"⊆": `\subseteq`,
	"⊇": `\supseteq`,
	"∈": `\in`,
	"∉": `\notin`,
	"∋": `\ni`,
	"→": `\rightarrow`,
	"←": `\leftarrow`,
	"↔": `\leftrightarrow`,
	"⇒": `\Rightarrow`,
	"⇐": `\Leftarrow`,
	"⇔": `\Leftrightarrow`,
	"↑": `\uparrow`,
	"↓": `\downarrow`,
	"↦": `\mapsto`,
	"∞": `\infty`,
	"∂": `\partial`,
	"∇": `\nabla`,
	"ℏ": `\hbar`,
	"∀": `\forall`,
	"∃": `\exists`,
	"∄": `\nexists`,
	"∅": `\emptyset`,
	"¬": `\neg`,
	"∧": `\land`,
	"∨": `\lor`,
	"∩": `\cap`,
	"∪": `\cup`,
	"∫": `\int`,
	"∬": `\iint`,
	"∭": `\iiint`,
	"∮": `\oint`,
	"∑": `\sum`,
	"∏": `\prod`,
	"∐": `\coprod`,
	"⟨": `\langle`,
	"⟩": `\rangle`,
	"⌈": `\lceil`,
	"⌉": `\rceil`,
	"⌊": `\lfloor`,
	"⌋": `\rfloor`,
	"…": `\ldots`,
	"⋯": `\cdots`,
	"⋮": `\vdots`,
	"⋱": `\ddots`,
}

// ParseOfficeMath 解析 m:oMath 元素。
func ParseOfficeMath(decoder *xml.Decoder, start xml.StartElement) (*OfficeMath, error) {
	content, err := parseMathContent(decoder, start.Name.Local)
	if err != nil {
		return nil, err
	}

	return &OfficeMath{
		XMLName: start.Name,
		Content: content,
	}, nil
}

// ParseOfficeMathPara 解析 m:oMathPara 元素。
func ParseOfficeMathPara(decoder *xml.Decoder, start xml.StartElement) (*OfficeMathPara, error) {
	para := &OfficeMathPara{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "oMath":
				math, err := ParseOfficeMath(decoder, t)
				if err != nil {
					return nil, err
				}
				para.Math = math
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return para, nil
			}
		}
	}
}

// OMMLToLaTeX 将 Office Math 结构转换为 LaTeX 字符串。
func OMMLToLaTeX(math *OfficeMath) string {
	if math == nil {
		return ""
	}
	return strings.TrimSpace(renderMathContent(math.Content))
}

// OMMLParaToLaTeX 将块级 Office Math 结构转换为 LaTeX 字符串。
func OMMLParaToLaTeX(para *OfficeMathPara) string {
	if para == nil {
		return ""
	}
	return OMMLToLaTeX(para.Math)
}

func parseMathContent(decoder *xml.Decoder, endLocal string) ([]interface{}, error) {
	content := make([]interface{}, 0)

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			element, err := parseMathElement(decoder, t)
			if err != nil {
				return nil, err
			}
			if element != nil {
				content = append(content, element)
			}
		case xml.EndElement:
			if t.Name.Local == endLocal {
				return content, nil
			}
		}
	}
}

func parseMathElement(decoder *xml.Decoder, start xml.StartElement) (interface{}, error) {
	switch start.Name.Local {
	case "r":
		return parseMathRun(decoder, start)
	case "f":
		return parseMathFrac(decoder, start)
	case "m":
		return parseMathMatrix(decoder, start)
	case "sSup":
		return parseMathSup(decoder, start)
	case "sSub":
		return parseMathSub(decoder, start)
	case "sSubSup":
		return parseMathSubSup(decoder, start)
	case "rad":
		return parseMathRad(decoder, start)
	case "d":
		return parseMathDelim(decoder, start)
	case "oMath":
		return ParseOfficeMath(decoder, start)
	default:
		if err := skipMathElement(decoder, start.Name.Local); err != nil {
			return nil, err
		}
		return nil, nil
	}
}

func parseMathRun(decoder *xml.Decoder, start xml.StartElement) (*MathRun, error) {
	run := &MathRun{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "t":
				text, err := readMathText(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				run.Text = &MathText{XMLName: t.Name, Content: text}
			case "rPr":
				runPr, err := parseMathRunProperties(decoder, t)
				if err != nil {
					return nil, err
				}
				run.RunPr = runPr
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return run, nil
			}
		}
	}
}

func parseMathRunProperties(decoder *xml.Decoder, start xml.StartElement) (*MathRunProp, error) {
	props := &MathRunProp{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "sty":
				props.Sty = &MathSty{
					XMLName: t.Name,
					Val:     getMathAttr(t.Attr, "val"),
				}
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return props, nil
			}
		}
	}
}

func parseMathFrac(decoder *xml.Decoder, start xml.StartElement) (*MathFrac, error) {
	frac := &MathFrac{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "fPr":
				fracPr, err := parseMathFracPr(decoder, t)
				if err != nil {
					return nil, err
				}
				frac.FracPr = fracPr
			case "num":
				content, err := parseMathContent(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				frac.Num = &MathNum{XMLName: t.Name, Content: content}
			case "den":
				content, err := parseMathContent(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				frac.Den = &MathDen{XMLName: t.Name, Content: content}
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return frac, nil
			}
		}
	}
}

func parseMathFracPr(decoder *xml.Decoder, start xml.StartElement) (*MathFracPr, error) {
	pr := &MathFracPr{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "type":
				pr.Type = &MathFracType{
					XMLName: t.Name,
					Val:     getMathAttr(t.Attr, "val"),
				}
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return pr, nil
			}
		}
	}
}

func parseMathMatrix(decoder *xml.Decoder, start xml.StartElement) (*MathMatrix, error) {
	matrix := &MathMatrix{
		XMLName: start.Name,
		Rows:    make([]*MathMatrixRow, 0),
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "mPr":
				matrixPr, err := parseMathMatrixPr(decoder, t)
				if err != nil {
					return nil, err
				}
				matrix.MatrixPr = matrixPr
			case "mr":
				row, err := parseMathMatrixRow(decoder, t)
				if err != nil {
					return nil, err
				}
				matrix.Rows = append(matrix.Rows, row)
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return matrix, nil
			}
		}
	}
}

func parseMathMatrixPr(decoder *xml.Decoder, start xml.StartElement) (*MathMatrixPr, error) {
	pr := &MathMatrixPr{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "mcs":
				columns, err := parseMathMatrixColumns(decoder, t)
				if err != nil {
					return nil, err
				}
				pr.Columns = columns
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return pr, nil
			}
		}
	}
}

func parseMathMatrixColumns(decoder *xml.Decoder, start xml.StartElement) (*MathMatrixColumns, error) {
	columns := &MathMatrixColumns{
		XMLName: start.Name,
		Columns: make([]*MathMatrixColumn, 0),
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "mc":
				column, err := parseMathMatrixColumn(decoder, t)
				if err != nil {
					return nil, err
				}
				columns.Columns = append(columns.Columns, column)
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return columns, nil
			}
		}
	}
}

func parseMathMatrixColumn(decoder *xml.Decoder, start xml.StartElement) (*MathMatrixColumn, error) {
	column := &MathMatrixColumn{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "mcPr":
				props, err := parseMathMatrixColumnPr(decoder, t)
				if err != nil {
					return nil, err
				}
				column.Props = props
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return column, nil
			}
		}
	}
}

func parseMathMatrixColumnPr(decoder *xml.Decoder, start xml.StartElement) (*MathMatrixColumnPr, error) {
	props := &MathMatrixColumnPr{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "count":
				props.Count = &MathCount{
					XMLName: t.Name,
					Val:     getMathAttr(t.Attr, "val"),
				}
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "mcJc":
				props.Jc = &MathMatrixColumnJustification{
					XMLName: t.Name,
					Val:     getMathAttr(t.Attr, "val"),
				}
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return props, nil
			}
		}
	}
}

func parseMathMatrixRow(decoder *xml.Decoder, start xml.StartElement) (*MathMatrixRow, error) {
	row := &MathMatrixRow{
		XMLName: start.Name,
		Cells:   make([]*MathE, 0),
	}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "e":
				content, err := parseMathContent(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				row.Cells = append(row.Cells, &MathE{XMLName: t.Name, Content: content})
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return row, nil
			}
		}
	}
}

func parseMathSup(decoder *xml.Decoder, start xml.StartElement) (*MathSup, error) {
	sup := &MathSup{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "e":
				content, err := parseMathContent(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				sup.E = &MathE{XMLName: t.Name, Content: content}
			case "sup":
				content, err := parseMathContent(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				sup.Sup = &MathSupElement{XMLName: t.Name, Content: content}
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return sup, nil
			}
		}
	}
}

func parseMathSub(decoder *xml.Decoder, start xml.StartElement) (*MathSub, error) {
	sub := &MathSub{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "e":
				content, err := parseMathContent(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				sub.E = &MathE{XMLName: t.Name, Content: content}
			case "sub":
				content, err := parseMathContent(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				sub.Sub = &MathSubElement{XMLName: t.Name, Content: content}
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return sub, nil
			}
		}
	}
}

func parseMathSubSup(decoder *xml.Decoder, start xml.StartElement) (*MathSubSup, error) {
	subSup := &MathSubSup{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "e":
				content, err := parseMathContent(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				subSup.E = &MathE{XMLName: t.Name, Content: content}
			case "sub":
				content, err := parseMathContent(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				subSup.Sub = &MathSubElement{XMLName: t.Name, Content: content}
			case "sup":
				content, err := parseMathContent(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				subSup.Sup = &MathSupElement{XMLName: t.Name, Content: content}
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return subSup, nil
			}
		}
	}
}

func parseMathRad(decoder *xml.Decoder, start xml.StartElement) (*MathRad, error) {
	rad := &MathRad{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "radPr":
				radPr, err := parseMathRadPr(decoder, t)
				if err != nil {
					return nil, err
				}
				rad.RadPr = radPr
			case "deg":
				content, err := parseMathContent(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				rad.Deg = &MathDeg{XMLName: t.Name, Content: content}
			case "e":
				content, err := parseMathContent(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				rad.E = &MathE{XMLName: t.Name, Content: content}
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return rad, nil
			}
		}
	}
}

func parseMathRadPr(decoder *xml.Decoder, start xml.StartElement) (*MathRadPr, error) {
	pr := &MathRadPr{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "degHide":
				pr.DegHide = &MathDegHide{
					XMLName: t.Name,
					Val:     getMathAttr(t.Attr, "val"),
				}
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return pr, nil
			}
		}
	}
}

func parseMathDelim(decoder *xml.Decoder, start xml.StartElement) (*MathDelim, error) {
	delim := &MathDelim{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "dPr":
				pr, err := parseMathDelimPr(decoder, t)
				if err != nil {
					return nil, err
				}
				delim.DPr = pr
			case "e":
				content, err := parseMathContent(decoder, t.Name.Local)
				if err != nil {
					return nil, err
				}
				delim.E = &MathE{XMLName: t.Name, Content: content}
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return delim, nil
			}
		}
	}
}

func parseMathDelimPr(decoder *xml.Decoder, start xml.StartElement) (*MathDelimPr, error) {
	pr := &MathDelimPr{XMLName: start.Name}

	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "begChr":
				pr.BegChr = &MathDelimChar{
					Val: getMathAttr(t.Attr, "val"),
				}
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			case "endChr":
				pr.EndChr = &MathDelimChar{
					Val: getMathAttr(t.Attr, "val"),
				}
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			default:
				if err := skipMathElement(decoder, t.Name.Local); err != nil {
					return nil, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == start.Name.Local {
				return pr, nil
			}
		}
	}
}

func renderMathContent(content []interface{}) string {
	var result strings.Builder

	for _, item := range content {
		result.WriteString(renderMathElement(item))
	}

	return result.String()
}

func renderMathElement(element interface{}) string {
	switch v := element.(type) {
	case *OfficeMath:
		return OMMLToLaTeX(v)
	case *MathRun:
		return renderMathRun(v)
	case *MathFrac:
		if v == nil || v.Num == nil || v.Den == nil {
			return ""
		}
		return `\frac{` + renderMathContent(v.Num.Content) + `}{` + renderMathContent(v.Den.Content) + `}`
	case *MathMatrix:
		return renderMathMatrix(v)
	case *MathSup:
		if v == nil || v.E == nil || v.Sup == nil {
			return ""
		}
		return renderMathContent(v.E.Content) + `^{` + renderMathContent(v.Sup.Content) + `}`
	case *MathSub:
		if v == nil || v.E == nil || v.Sub == nil {
			return ""
		}
		return renderMathContent(v.E.Content) + `_{` + renderMathContent(v.Sub.Content) + `}`
	case *MathSubSup:
		if v == nil || v.E == nil || v.Sub == nil || v.Sup == nil {
			return ""
		}
		return renderMathContent(v.E.Content) + `_{` + renderMathContent(v.Sub.Content) + `}^{` + renderMathContent(v.Sup.Content) + `}`
	case *MathRad:
		if v == nil || v.E == nil {
			return ""
		}
		if v.Deg != nil && len(v.Deg.Content) > 0 {
			return `\sqrt[` + renderMathContent(v.Deg.Content) + `]{` + renderMathContent(v.E.Content) + `}`
		}
		return `\sqrt{` + renderMathContent(v.E.Content) + `}`
	case *MathDelim:
		if v == nil || v.E == nil {
			return ""
		}
		if matrixEnv, ok := renderDelimitedMatrix(v); ok {
			return matrixEnv
		}
		begin := "("
		end := ")"
		if v.DPr != nil && v.DPr.BegChr != nil && v.DPr.BegChr.Val != "" {
			begin = v.DPr.BegChr.Val
		}
		if v.DPr != nil && v.DPr.EndChr != nil && v.DPr.EndChr.Val != "" {
			end = v.DPr.EndChr.Val
		}
		return begin + renderMathContent(v.E.Content) + end
	default:
		return ""
	}
}

func renderMathMatrix(matrix *MathMatrix) string {
	if matrix == nil {
		return ""
	}

	rows := make([]string, 0, len(matrix.Rows))
	for _, row := range matrix.Rows {
		if row == nil {
			continue
		}

		cells := make([]string, 0, len(row.Cells))
		for _, cell := range row.Cells {
			if cell == nil {
				cells = append(cells, "")
				continue
			}
			cells = append(cells, renderMathContent(cell.Content))
		}
		rows = append(rows, strings.Join(cells, " & "))
	}

	return `\begin{matrix}` + strings.Join(rows, ` \\ `) + `\end{matrix}`
}

func renderDelimitedMatrix(delim *MathDelim) (string, bool) {
	if delim == nil || delim.E == nil || len(delim.E.Content) != 1 {
		return "", false
	}

	matrix, ok := delim.E.Content[0].(*MathMatrix)
	if !ok {
		return "", false
	}

	env := matrixEnvironmentFromDelim(delim)
	if env == "" {
		return "", false
	}

	matrixBody := strings.TrimPrefix(renderMathMatrix(matrix), `\begin{matrix}`)
	matrixBody = strings.TrimSuffix(matrixBody, `\end{matrix}`)
	return `\begin{` + env + `}` + matrixBody + `\end{` + env + `}`, true
}

func matrixEnvironmentFromDelim(delim *MathDelim) string {
	if delim == nil || delim.DPr == nil || delim.DPr.BegChr == nil || delim.DPr.EndChr == nil {
		return ""
	}

	begin := delim.DPr.BegChr.Val
	end := delim.DPr.EndChr.Val
	switch {
	case begin == "[" && end == "]":
		return "bmatrix"
	case begin == "(" && end == ")":
		return "pmatrix"
	case begin == "|" && end == "|":
		return "vmatrix"
	case begin == "{" && end == "}":
		return "Bmatrix"
	default:
		return ""
	}
}

func renderMathRun(run *MathRun) string {
	if run == nil || run.Text == nil {
		return ""
	}

	var result strings.Builder
	for _, r := range run.Text.Content {
		token := string(r)
		if latex, ok := unicodeToLaTeXToken[token]; ok {
			result.WriteString(latex)
			continue
		}
		result.WriteRune(r)
	}

	return result.String()
}

func getMathAttr(attrs []xml.Attr, name string) string {
	for _, attr := range attrs {
		if attr.Name.Local == name {
			return attr.Value
		}
	}
	return ""
}

func readMathText(decoder *xml.Decoder, endLocal string) (string, error) {
	var result strings.Builder

	for {
		token, err := decoder.Token()
		if err != nil {
			return "", err
		}

		switch t := token.(type) {
		case xml.CharData:
			result.Write([]byte(t))
		case xml.EndElement:
			if t.Name.Local == endLocal {
				return result.String(), nil
			}
		}
	}
}

func skipMathElement(decoder *xml.Decoder, elementName string) error {
	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			return err
		}

		switch t := token.(type) {
		case xml.StartElement:
			if t.Name.Local == elementName {
				depth++
			}
		case xml.EndElement:
			if t.Name.Local == elementName {
				depth--
			}
		}
	}

	return nil
}
