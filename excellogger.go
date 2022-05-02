package excelutils

import (
	"errors"
	"os"
	"time"

	"github.com/xuri/excelize/v2"
)

var line, col, auforFilterStartcol, auforFilterStartrow, maxcol int
var fexcel *excelize.File
var sheet string

func IsWritable(name string) (isWritable bool) {
	isWritable = false

	_, err := os.OpenFile(name, os.O_WRONLY|os.O_TRUNC|os.O_CREATE, 0666)
	if err != nil {
		return isWritable
	}
	isWritable = true
	return isWritable
}

func Check(e error) {
	if e != nil {
		panic(e)
	}
}

func max(x, y int) int {
	if x < y {
		return y
	}
	return x
}

type Options struct {
	SheetName string
}

func NewFile(opt *Options) {
	fexcel = excelize.NewFile()
	line = 1
	col = 1
	maxcol = 1
	if opt != nil {
		sheet = opt.SheetName
	} else {
		sheet = "Sheet1"
	}
	fexcel.WorkBook.Sheets.Sheet[0].Name = sheet
}

func OpenFile(filename string) (err error) {
	fexcel, err = excelize.OpenFile(filename)
	line = 1
	col = 1
	maxcol = 1
	sheet = fexcel.WorkBook.Sheets.Sheet[0].Name
	return err
}

func NewSheet(name string) error {

	for _, oldsheets := range fexcel.WorkBook.Sheets.Sheet {
		if oldsheets.Name == name {
			return errors.New("sheet already exists")
		}
	}
	fexcel.NewSheet(name)
	sheet = name
	return nil
}

func SelectSheet(name string) error {

	for _, oldsheets := range fexcel.WorkBook.Sheets.Sheet {
		if oldsheets.Name == name {
			sheet = oldsheets.Name
			line = 1
			col = 1
			maxcol = 1
			return nil
		}
	}
	return errors.New("sheet not found")
}

func NextLine() {
	line++
	ResetCol()
}

func NextCol() {
	col++
	maxcol = max(maxcol, col)
}

func ResetCol() {
	col = 1
}

func WriteColumnsHeaderln(data []string) {
	for _, v := range data {
		SetTableHeader()
		WiteCell(v)
		col++
		maxcol = max(maxcol, col)
	}
	col = 1
	line++
}

func WriteColumnsHeaderRotln(data []string) {
	for _, v := range data {
		SetTableHeaderRot()
		WiteCell(v)
		col++
		maxcol = max(maxcol, col)
	}
	col = 1
	line++
}

func WriteColumns(data []string) {
	for _, v := range data {
		WiteCell(v)
		col++
		maxcol = max(maxcol, col)
	}
}

func WriteColumnsln(data []string) {
	WriteColumns(data)
	col = 1
	line++
}

func WiteCell(msg interface{}) {
	axis, err := excelize.CoordinatesToCellName(col, line)
	Check(err)
	err = fexcel.SetCellValue(sheet, axis, msg)
	Check(err)
}

func WiteCellln(msg interface{}) {
	WiteCell(msg)
	line++
	col = 1
}

func WiteCellnc(msg interface{}) {
	WiteCell(msg)
	col++
	maxcol = max(maxcol, col)
}

func WiteCellHyperLinknc(msg interface{}, hyperlink string) {
	axis, err := excelize.CoordinatesToCellName(col, line)
	Check(err)
	err = fexcel.SetCellValue(sheet, axis, msg)
	Check(err)
	err = fexcel.SetCellHyperLink(sheet, axis, hyperlink, "External")
	Check(err)
	col++
	maxcol = max(maxcol, col)
}

func WiteBoolCellnc(msg bool) {
	if msg {
		WiteCell("X")
	} else {
		WiteCell("-")
	}
	col++
	maxcol = max(maxcol, col)
}

func SetCellStyleRotateXY(q, v int) {
	axis, err := excelize.CoordinatesToCellName(q, v)
	Check(err)
	style, err := fexcel.NewStyle(`{"alignment":{"text_rotation":90,"horizontal":"center"},"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":1}}`)
	Check(err)
	err = fexcel.SetCellStyle(sheet, axis, axis, style)
	Check(err)
}

func SetCellStyleCenter() {
	axis, err := excelize.CoordinatesToCellName(col, line)
	Check(err)
	//var style excelize.Style
	style, err2 := fexcel.NewStyle(`{"alignment":{"horizontal":"center"}}`)
	Check(err2)
	err = fexcel.SetCellStyle(sheet, axis, axis, style)
	Check(err)
}

func SetCellStyleColor(color string) {
	axis, err := excelize.CoordinatesToCellName(col, line)
	Check(err)
	style, err := fexcel.NewStyle(`{"fill":{"type":"pattern","color":["` + color + `"],"pattern":1}}`)
	Check(err)
	err = fexcel.SetCellStyle(sheet, axis, axis, style)
	Check(err)
}

func SetCellStyleRotate() {
	SetCellStyleRotateXY(col, line)
}

func SetCellStyleRotateN(count int) {
	var i, j int
	i = col
	j = line
	for n := 0; n < count; n++ {
		SetCellStyleRotateXY(i, j)
		i++
	}
}

func SetCellFontHeader() {
	axis, err := excelize.CoordinatesToCellName(col, line)
	Check(err)
	//var style excelize.Style
	style, err2 := fexcel.NewStyle(`{"font":{"bold":true,"family":"Times New Roman","size":24,"color":"#777777"}}`)
	Check(err2)
	err = fexcel.SetCellStyle(sheet, axis, axis, style)
	Check(err)
	err = fexcel.SetRowHeight(sheet, line, 24)
	Check(err)
}

func SetCellFontHeader2() {
	axis, err := excelize.CoordinatesToCellName(col, line)
	Check(err)
	//var style excelize.Style
	style, err2 := fexcel.NewStyle(`{"font":{"bold":true,"family":"Times New Roman","size":16,"color":"#777777"}}`)
	Check(err2)
	err = fexcel.SetCellStyle(sheet, axis, axis, style)
	Check(err)
	err = fexcel.SetRowHeight(sheet, line, 16)
	Check(err)
}

func SetTableHeader() {
	cellname, err := excelize.CoordinatesToCellName(col, line)
	Check(err)
	style, err := fexcel.NewStyle(`{"alignment":{"horizontal":"center","vertical":"center"},"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":1}}`)
	Check(err)
	err = fexcel.SetCellStyle(sheet, cellname, cellname, style)
	Check(err)
}

func SetTableHeaderRot() {
	cellname, err := excelize.CoordinatesToCellName(col, line)
	Check(err)
	style, err := fexcel.NewStyle(`{"alignment":{"text_rotation":90,"horizontal":"center","vertical":"center"},"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":1}}`)
	Check(err)
	err = fexcel.SetCellStyle(sheet, cellname, cellname, style)
	Check(err)
}

func AutoFilterStart() {
	auforFilterStartcol = col
	auforFilterStartrow = line
}

func AutoFilterEnd() {
	axis, err := excelize.CoordinatesToCellName(auforFilterStartcol, auforFilterStartrow)
	Check(err)
	autoFilter(axis)
}

func autoFilter(uppperleft string) {
	rows, err := fexcel.GetRows(sheet)
	Check(err)
	nrows := len(rows)
	ncols := maxcol
	axis, err2 := excelize.CoordinatesToCellName(ncols, nrows)
	Check(err2)
	err2 = fexcel.AutoFilter(sheet, uppperleft, axis, "")
	Check(err2)
}

func SetCell(msg interface{}, x int, y int) {
	axis, err := excelize.CoordinatesToCellName(x, y)
	Check(err)
	err = fexcel.SetCellValue(sheet, axis, msg)
	Check(err)
}

func SetCellBackgroundAxis(axis, color string) {
	style, err := fexcel.NewStyle(`{"fill":{"type":"pattern","color":["` + color + `"],"pattern":1}}`)
	Check(err)
	err = fexcel.SetCellStyle(sheet, axis, axis, style)
	Check(err)
}

func SetCellBackground(color string, x int, y int) {
	axis, err := excelize.CoordinatesToCellName(x, y)
	Check(err)
	SetCellBackgroundAxis(axis, color)
}

func SetColWidth(startcol, endcol string, width float64) {
	err := fexcel.SetColWidth(sheet, startcol, endcol, width)
	Check(err)
}

func SetAutoColWidth() {
	rows, err := fexcel.GetRows(sheet)
	Check(err)

	for k := 1; k < maxcol+1; k++ {
		maxwidth := 0
		for row := range rows {
			axis, err2 := excelize.CoordinatesToCellName(k, row+1)
			Check(err2)
			strlen, _ := fexcel.GetCellValue(sheet, axis)
			maxwidth = max(maxwidth, len(strlen))
		}
		colname, _ := excelize.ColumnNumberToName(k)
		width := maxwidth + 5
		if width > 200 {
			width = 200
		}
		err = fexcel.SetColWidth(sheet, colname, colname, float64(width))
		Check(err)
	}
}

func SetRowHeight(height float64) {
	err := fexcel.SetRowHeight(sheet, line, height)
	Check(err)
}

func SaveAs(name string) {
	if !IsWritable(name) {
		time.Sleep(time.Second)
	}
	err := fexcel.SaveAs(name)
	Check(err)
}

func BoolToEmoji(syn bool) string {
	emo := ""
	if syn {
		emo = "✔"
	} else {
		emo = "❌"
	}
	return emo
}
