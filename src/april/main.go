package main

import (
	"april/model"
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"bufio"
	"path/filepath"
)

func main() {
	if len(os.Args) < 4 || os.Args[1] != "CAT5" {
		fmt.Printf("usage: %s CAT5 invoiceFile xslFile\n", os.Args[0])
		return
	}
	invoiceFileName := os.Args[2]
	xslFileName := os.Args[3]

	invoiceFile, err := os.Open(invoiceFileName)
	if err != nil {
		fmt.Printf("open invoice file %s failed: %s\n", invoiceFile, err)
		return
	}
	defer invoiceFile.Close()

	xslFile, err := xlsx.OpenFile(xslFileName)
	if err != nil {
		fmt.Printf("open file %s failed: %s\n", xslFileName, err)
		return
	}

	reader := bufio.NewReader(invoiceFile)
	invoices := make(map[string]bool)
	for {
		invoiceNo := ""
		if _, err := fmt.Fscanf(reader, "%s\r\n", &invoiceNo); err != nil {
			break
		}
		fmt.Printf("found invoice No: %s\n", invoiceNo)
		invoices[invoiceNo] = true
	}

	if err := model.CAT5(invoices, xslFile); err != nil {
		fmt.Printf("CAT5 failed: %s\n", xslFile)
		return
	}

	newFileName := filepath.Dir(xslFileName) + "\\NEW-" + filepath.Base(xslFileName)
	fmt.Printf("%s\n", newFileName)
	if err := xslFile.Save(newFileName); err != nil {
		fmt.Printf("save to file %s failed: %s\n", newFileName, err)
		return
	}
}
