package main

import (
	"fmt"

	"github.com/tealeg/xlsx"
)

func main() {
	excelFileName := "./data.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		panic(err)
	}

	sheet := xlFile.Sheets[0] // care only about first sheet
	dataSamples := make(map[string]int)
	fmt.Println("extract sample data..")
	for i, row := range sheet.Rows {
		for j, cell := range row.Cells {
			if j != 0 {
				break
			}
			text := cell.String()
			dataSamples[text] = i
		}
	}

	crcFile, err := xlsx.OpenFile("crc.xlsx")
	if err != nil {
		panic(err)
	}
	sheet2 := crcFile.Sheets[0] // care only about first sheet
	crcSamples := make(map[string]int)
	fmt.Println("extract crc sample data..")
	for i, row := range sheet2.Rows {
		for j, cell := range row.Cells {
			if j != 0 {
				break
			}
			text := cell.String()
			crcSamples[text] = i
		}
	}

	// get matched samples from dataSamples in crcSamples
	matched := make(map[string]int)
	fmt.Println("finding match with crc sample data..")
	for k, _ := range dataSamples {
		if _, ok := crcSamples[k]; ok {
			matched[k] = crcSamples[k]
		}
	}

	// we need 5..15 index inclusive
	srcRows := sheet2.Rows
	dstRows := sheet.Rows

	updatedRows := make([]*xlsx.Row, 0)

	fmt.Println("writing updated rows..")
	for rowText, rowIndex := range matched {
		srcRow := srcRows[rowIndex]

		dstRowIndex := dataSamples[rowText]

		dstRow := dstRows[dstRowIndex]

		// print dstRow

		writeRow(srcRow, dstRow)

		updatedRows = append(updatedRows, dstRow)
	}

	// for _, up := range updatedRows {
	// 	fmt.Println("len of dstRow", len(up.Cells))

	// }

	// printUpdatedRows(updatedRows)

	if err := xlFile.Save("updated-data.xlsx"); err != nil {
		panic(err)
	}
}

func writeRow(src, dst *xlsx.Row) {
	srcCells := src.Cells // 5-15
	dstCells := dst.Cells // 51-61

	if len(dstCells) > 51 {
		return // already vineetha did the job
	}

	// fmt.Println("dst cells len", len(dstCells))

	for i := 5; i <= 15; i++ {
		cell := xlsx.NewCell(dst)
		cell.SetString(srcCells[i].String())
		dst.Cells = append(dst.Cells, cell)
	}
}

func printUpdatedRows(rows []*xlsx.Row) {
	for _, row := range rows {
		cells := row.Cells
		for i := 51; i < 61; i++ {
			fmt.Printf("%s,", cells[i])
		}
		fmt.Println()
	}
}
