package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	f := excelize.NewFile()
	sheetName := "Transcript"
	f.SetSheetName("Sheet1", sheetName)

	data := [][]interface{}{
		{"Student Exame Score"},
		{"Type: Midterm Exam", nil, nil, nil, "Core Curriculum Science", nil, nil, "Science"},
		{"Number", "ID", "Name", "Class", "Language Arts", "Maths", "History", "Chemistry", "Biology", "Physics", "Total"},
		{1, 10001, "Student A", "Class 1", 93, 80, 89, 86, 57, 77},
		{2, 10002, "Student B", "Class 2", 65, 77, 19, 25, 54, 12},
	}
	for i, row := range data {
		startCell, err := excelize.JoinCellName("A", i+1)
		if err != nil {
			fmt.Println(err)
			return
		}
		if err := f.SetSheetRow(sheetName, startCell, &row); err != nil {
			fmt.Println(err)
			return
		}
	}
	formulaType, ref := excelize.STCellFormulaTypeShared, "K4:K9"
	if err := f.SetCellFormula(sheetName, "K4", "=SUM(E4:J4)",
		excelize.FormulaOpts{Ref: &ref, Type: &formulaType}); err != nil {
		fmt.Println(err)
		return
	}

	mergeCellRanges := [][]string{{"A1", "K1"}, {"A2", "D2"}, {"E2", "G2"}, {"H2", "K2"}}
	for _, ranges := range mergeCellRanges {
		if err := f.MergeCell(sheetName, ranges[0], ranges[1]); err != nil {
			fmt.Println(err)
			return
		}
	}
	style1, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#D07462"}, Pattern: 1},
	})

	if err != nil {
		fmt.Println(err)
		return
	}
	if f.SetCellStyle(sheetName, "A1", "A1", style1); err != nil {
		fmt.Println(err)
		return
	}

	style2, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#C9DFB9"}, Pattern: 1},
	})

	for _, cell := range []string{"A2", "E2", "H2"} {
		if err := f.SetCellStyle(sheetName, cell, cell, style2); err != nil {
			fmt.Println(err)
			return
		}
	}

	if err := f.SetColWidth(sheetName, "D", "K", 14); err != nil {
		fmt.Println(err)
		return
	}

	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}

}
