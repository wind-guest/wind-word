// Package markdown 保留数学公式类型与转换 API 的兼容导出。
package markdown

import "github.com/FengKeWG/wind-word/pkg/omml"

type OfficeMath = omml.OfficeMath
type OfficeMathPara = omml.OfficeMathPara
type MathRun = omml.MathRun
type MathText = omml.MathText
type MathRunProp = omml.MathRunProp
type MathSty = omml.MathSty
type MathFrac = omml.MathFrac
type MathFracPr = omml.MathFracPr
type MathFracType = omml.MathFracType
type MathNum = omml.MathNum
type MathDen = omml.MathDen
type MathSup = omml.MathSup
type MathE = omml.MathE
type MathSupElement = omml.MathSupElement
type MathSub = omml.MathSub
type MathSubElement = omml.MathSubElement
type MathRad = omml.MathRad
type MathRadPr = omml.MathRadPr
type MathDegHide = omml.MathDegHide
type MathDeg = omml.MathDeg
type MathSubSup = omml.MathSubSup
type MathDelim = omml.MathDelim
type MathDelimPr = omml.MathDelimPr
type MathDelimChar = omml.MathDelimChar

func LaTeXToOMML(latex string) *OfficeMath {
	return omml.LaTeXToOMML(latex)
}

func LaTeXToOMMLString(latex string, isBlock bool) (string, error) {
	return omml.LaTeXToOMMLString(latex, isBlock)
}
