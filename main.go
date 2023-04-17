package main

import (
	"math/rand"

	"github.com/xuri/excelize/v2"
)

func firstExample() {

	file := excelize.NewFile()

	index, err := file.NewSheet("sheet2")

	if err != nil {
		return
	}

	file.SetCellValue("sheet2", "A2", "Hello world.")

	file.SetCellValue("sheet1", "B2", 100)

	file.SetActiveSheet(index)

	err = file.SaveAs("test.xlsx")

	if err != nil {
		return
	}

}

func main() {

	file := excelize.NewFile()

	sw, err := file.NewStreamWriter("sheet1")

	checkError(err)

	styleID, err := file.NewStyle(&excelize.Style{Font: &excelize.Font{Color: "7777777"}})

	checkError(err)

	err = sw.SetRow("A1", []interface{}{excelize.Cell{StyleID: styleID, Value: "Data"},
		[]excelize.RichTextRun{
			{Text: "Rich", Font: &excelize.Font{Color: "2354e8"}},
			{Text: "Text", Font: &excelize.Font{Color: "e83723"}},
		}}, excelize.RowOpts{Height: 45, Hidden: false})

	checkError(err)

	for rowId := 2; rowId <= 102400; rowId++ {
		row := make([]interface{}, 50)

		for colId := 0; colId < 50; colId++ {
			row[colId] = rand.Intn(6400000)
		}
		cell, err := excelize.CoordinatesToCellName(1, rowId)
		checkError(err)
		sw.SetRow(cell, row)
	}

	sw.Flush()

	file.SaveAs("Streamed.xlsx")

}

func checkError(err error) {

	if err != nil {
		panic(err)
	}
}
