package model

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"time"
	"april/util"
)

func Test(file *xlsx.File) error {
	sheet := file.Sheets[0]
	for rowNum, row := range sheet.Rows {
		fmt.Printf("row %d:\t", rowNum)
		for _, cell := range row.Cells {
			text, _ := cell.String()
			fmt.Printf("%s\t", text)
		}
		fmt.Printf("\n")
		if rowNum > 20 {
			break
		}
	}
	return nil
}

func CAT5(invoices map[string]bool, file *xlsx.File) error {
	sheet := file.Sheets[0]
	for rowNum, row := range sheet.Rows {
		if (rowNum < 14) {
			continue
		}

		if len(row.Cells) < 70 {
			break
		}

		accountNo, err := row.Cells[1].String()
		if err != nil {
			fmt.Printf("get value of row %d cell 5 failed: %s\n", rowNum, err)
			continue
		}

		if _, exists := invoices[accountNo]; !exists {
			continue
		}

		row.Cells[34].SetString("Finalized")
		row.Cells[35].SetString("Finalized")
		row.Cells[38].SetDateTimeWithFormat(util.TimeToExcelTime(time.Now().UTC()), "mm/dd/yyyy")
		row.Cells[62].SetDateTimeWithFormat(util.TimeToExcelTime(time.Now().UTC()), "mm/dd/yyyy")
		row.Cells[61].SetString("CTA5")
		if str, err := row.Cells[22].String(); err == nil && str == "In process of WHT and VAT filing" {
			row.Cells[22].SetString("")
		}
		if str, err := row.Cells[69].String(); err == nil && str == "In process of WHT and VAT filing" {
			row.Cells[69].SetString("")
		}
	}
	return nil
}
