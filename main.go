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

	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}

}
