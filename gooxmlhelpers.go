package gooxmlhelpers // import "github.com/l0rda/gooxmlhelpers"

import (
	"strings"

	"baliance.com/gooxml/color"
	"baliance.com/gooxml/schema/soo/sml"
	"baliance.com/gooxml/spreadsheet"
)

// Formula originally was posted at http://www.excelworld.ru/forum/3-9902-1
// by MCH (http://www.excelworld.ru/index/8-41)
const excelNumToTextFormula = `=SUBSTITUTE(PROPER(INDEX(n_4,MID(TEXT(A1,n0),1,1)+1)&INDEX(n0x,MID(TEXT(A1,n0),2,1)+1,MID(TEXT(A1,n0),3,1)+1)&IF(-MID(TEXT(A1,n0),1,3),"миллиард"&VLOOKUP(MID(TEXT(A1,n0),3,1)*AND(MID(TEXT(A1,n0),2,1)-1),mil,2),"")&INDEX(n_4,MID(TEXT(A1,n0),4,1)+1)&INDEX(n0x,MID(TEXT(A1,n0),5,1)+1,MID(TEXT(A1,n0),6,1)+1)&IF(-MID(TEXT(A1,n0),4,3),"миллион"&VLOOKUP(MID(TEXT(A1,n0),6,1)*AND(MID(TEXT(A1,n0),5,1)-1),mil,2),"")&INDEX(n_4,MID(TEXT(A1,n0),7,1)+1)&INDEX(n1x,MID(TEXT(A1,n0),8,1)+1,MID(TEXT(A1,n0),9,1)+1)&IF(-MID(TEXT(A1,n0),7,3),VLOOKUP(MID(TEXT(A1,n0),9,1)*AND(MID(TEXT(A1,n0),8,1)-1),ths,2),"")&INDEX(n_4,MID(TEXT(A1,n0),10,1)+1)&INDEX(n0x,MID(TEXT(A1,n0),11,1)+1,MID(TEXT(A1,n0),12,1)+1)),"z"," ")&IF(TRUNC(TEXT(A1,n0)),"","Ноль ")&"рубл"&VLOOKUP(MOD(MAX(MOD(MID(TEXT(A1,n0),11,2)-11,100),9),10),{0,"ь ";1,"я ";4,"ей "},2)&RIGHT(TEXT(A1,n0),2)&" копе"&VLOOKUP(MOD(MAX(MOD(RIGHT(TEXT(A1,n0),2)-11,100),9),10),{0,"йка";1,"йки";4,"ек"},2)`

var defNames = map[string]string{
	"n_1": `{"","одинz","дваz","триz","четыреz","пятьz","шестьz","семьz","восемьz","девятьz"}`,
	"n_2": `{"десятьz","одиннадцатьz","двенадцатьz","тринадцатьz","четырнадцатьz","пятнадцатьz","шестнадцатьz","семнадцатьz","восемнадцатьz","девятнадцатьz"}`,
	"n_3": `{"";1;"двадцатьz";"тридцатьz";"сорокz";"пятьдесятz";"шестьдесятz";"семьдесятz";"восемьдесятz";"девяностоz"}`,
	"n_4": `{"","стоz","двестиz","тристаz","четырестаz","пятьсотz","шестьсотz","семьсотz","восемьсотz","девятьсотz"}`,
	"n_5": `{"","однаz","двеz","триz","четыреz","пятьz","шестьz","семьz","восемьz","девятьz"}`,
	"n0":  `"000000000000"&MID(1/2,2,1)&"00"`,
	"n0x": `IF(n_3=1,n_2,n_3&n_1)`,
	"n1x": `IF(n_3=1,n_2,n_3&n_5)`,
	"mil": `{0,"овz";1,"z";2,"аz";5,"овz"}`,
	"ths": `{0,"тысячz";1,"тысячаz";2,"тысячиz";5,"тысячz"}`,
}

// SetDefinedNamesRub - set defined names for SetNum2RubFormula
func SetDefinedNamesRub(wb *spreadsheet.Workbook) {
	for k, v := range defNames {
		wb.AddDefinedName(k, v)
	}
}

// GetSpellFormula - return num2spell formula for cell (Russian rubles)
func GetSpellFormula(ref string) string {
	return strings.Replace(excelNumToTextFormula, "A1", ref, 25)
}

// SetSpellFormula - convert ref cell value(number) to words (Russian rubles), you need to run SetDefinedNamesRub()
func SetSpellFormula(cell spreadsheet.Cell, ref string) {
	cell.SetFormulaRaw(GetSpellFormula(ref))
}

// FillColor - fill cell by reference with color and save current cell style
func FillColor(ss spreadsheet.StyleSheet, cell spreadsheet.Cell, clr color.Color) {
	origStylePointer := cell.X().SAttr
	cs := ss.AddCellStyle()
	fills := ss.Fills()
	f := fills.AddFill()
	f.SetPatternFill().SetFgColor(clr)
	if origStylePointer == nil {
		// don't need to copy style
		cs.SetFill(f)
		cell.SetStyle(cs)
		return
	}
	var copySrc *sml.CT_Xf
	for i, doc := range ss.X().CellXfs.Xf {
		if uint32(i) == *origStylePointer {
			copySrc = doc
		}
		// new index always after source index in slice
		if uint32(i) == cs.Index() {
			*doc = *copySrc
			id := f.Index()
			doc.FillIdAttr = &id
			break
		}
	}
	cell.SetStyle(cs)
}

// SetNumberFormat - set number format to cell by reference and save current cell style
func SetNumberFormat(ss spreadsheet.StyleSheet, cell spreadsheet.Cell, format string) {
	// TODO: validate format
	origStylePointer := cell.X().SAttr
	cs := ss.AddCellStyle()
	cs.SetNumberFormat(format)
	if origStylePointer == nil {
		// don't need to copy style
		cell.SetStyle(cs)
		return
	}
	var copySrc *sml.CT_Xf
	for i, doc := range ss.X().CellXfs.Xf {
		if uint32(i) == *origStylePointer {
			copySrc = doc
		}
		// new index always after source index in slice
		if uint32(i) == cs.Index() {
			*doc = *copySrc
			id := cs.NumberFormat()
			doc.NumFmtIdAttr = &id
			break
		}
	}
	cell.SetStyle(cs)
}
