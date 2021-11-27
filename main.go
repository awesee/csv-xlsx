package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"runtime"
	"strings"

	"github.com/xuri/excelize/v2"
)

const VERSION = "1.0.0"

func main() {
	args := os.Args
	cmd := filepath.Base(args[0])
	if len(args) < 2 {
		fmt.Printf("%s version %s %s/%s\n\n", cmd, VERSION, runtime.GOOS, runtime.GOARCH)
		fmt.Printf("usage:\n\t%s file.csv|file.xlsx...\n", cmd)
		return
	}
	for _, filename := range args[1:] {
		src, err := os.Open(filename)
		check(err)
		ext := filepath.Ext(filename)
		name := strings.TrimSuffix(filename, ext)
		basename := filepath.Base(filename)
		switch strings.ToLower(ext) {
		case ".csv":
			csvToXlsx(src, name)
		case ".xlsx":
			xlsxToCsv(src, name)
		default:
			fmt.Printf("skipping file %s\n", basename)
		}
		_ = src.Close()
	}
}

func csvToXlsx(src io.Reader, name string) {
	csvFile := csv.NewReader(src)
	csvFile.FieldsPerRecord = -1
	xlsxFile := excelize.NewFile()
	maxRows := excelize.TotalRows
	var streamWriter *excelize.StreamWriter
	streamWriterFlush := func() {
		if streamWriter != nil {
			check(streamWriter.Flush())
		}
	}
	for rowID := 0; ; rowID++ {
		record, err := csvFile.Read()
		if err == io.EOF {
			break
		}
		check(err)
		if rowID%maxRows == 0 {
			streamWriterFlush()
			sheet := fmt.Sprintf("Sheet%d", rowID/maxRows+1)
			xlsxFile.NewSheet(sheet)
			streamWriter, err = xlsxFile.NewStreamWriter(sheet)
			check(err)
		}
		row := make([]interface{}, len(record))
		for k, v := range record {
			row[k] = v
		}
		axis, err := excelize.CoordinatesToCellName(1, rowID%maxRows+1)
		check(err)
		check(streamWriter.SetRow(axis, row))
	}
	streamWriterFlush()
	check(xlsxFile.SaveAs(name + ".xlsx"))
}

func xlsxToCsv(src io.Reader, name string) {
	xlsxFile, err := excelize.OpenReader(src)
	check(err)
	for _, sheet := range xlsxFile.GetSheetList() {
		filename := name
		if sheet != "Sheet1" {
			filename += "_" + sheet
		}
		csvFile, err := os.Create(filename + ".csv")
		check(err)
		csvWriter := csv.NewWriter(csvFile)
		rows, err := xlsxFile.Rows(sheet)
		check(err)
		totalCols := 0
		for rows.Next() {
			row, err := rows.Columns()
			check(err)
			if totalCols < 1 {
				totalCols = len(row)
			}
			for i := totalCols - len(row); i > 0; i-- {
				row = append(row, "")
			}
			check(csvWriter.Write(row))
		}
		csvWriter.Flush()
	}
}

func check(err error) {
	if err != nil {
		log.Fatal(err)
	}
}
