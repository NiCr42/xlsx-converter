package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"log"
	"os"

	"github.com/tealeg/xlsx"
)

// https://gist.github.com/jmoiron/e9f72720cef51862b967
// https://play.golang.org/p/FqKzq_1ICs
type Options struct {
	OutputFile string
	SheetIndex int
	SheetName  string
	HeaderLine int
	StartLine  int
	EndLine    int
	Limit      int
}

var options *Options

func init() {
	// Init options
	options = &Options{}

	// Custom usage display
	flag.Usage = func() {
		fmt.Fprintf(flag.CommandLine.Output(), "Print a sheet from a XLSX file to standard ouput (Stdout) as CSV format.\n")
		fmt.Fprintf(flag.CommandLine.Output(), "Usage: [-output-file <file> (-sheet-index <index> | -sheet-name <name>) -header-line <line> -start-line <line> -limit <limit-number-of-lines>] <file>\n")
		flag.PrintDefaults()
	}

	// Let the flag package handle options
	flag.StringVar(&options.OutputFile, "output-file", "", "CSV output file.")
	flag.IntVar(&options.SheetIndex, "sheet-index", 0, "Index of worksheet.")
	flag.StringVar(&options.SheetName, "sheet-name", "", "Name of worksheet.")
	flag.IntVar(&options.HeaderLine, "header-line", -1, "Index of header line.")
	flag.IntVar(&options.StartLine, "start-line", -1, "Index of start line.")
	flag.IntVar(&options.Limit, "limit", 0, "Limit number of lines to retrieve.")
	flag.Parse()

	if len(flag.Args()) < 1 {
		flag.Usage()
		os.Exit(-1)
	}
}

func main() {
	// Open XLSX file
	xlFile, err := xlsx.OpenFile(flag.Arg(0))
	if err != nil {
		log.Fatal("Error: ", err)
	}

	// Retrieve sheet
	sheet, err := setSheet(xlFile)
	if err != nil {
		log.Fatal("Error: ", err)
	}

	// Retrieve rows from sheet
	rows, err := setRows(sheet)
	if err != nil {
		log.Fatal("Error: ", err)
	}

	// Set io.Writer
	var writer io.Writer

	// If options.OutputFile was provided, lets try to create a file
	// Else output CSV to Stdout.
	if options.OutputFile != "" {
		writer, err := os.Create(options.OutputFile)
		defer writer.Close()

		if err != nil {
			log.Fatal("Error: ", err)
		}
	} else {
		writer = os.Stdout
	}

	// Initialize a CSV NewWriter with defined writer
	csv := csv.NewWriter(writer)
	defer csv.Flush()

	// Write file/Display to Stdout
	outputCsv(csv, rows)
}

func setSheet(xlFile *xlsx.File) (*xlsx.Sheet, error) {
	var sheet *xlsx.Sheet

	if options.SheetName != "" {
		if _, exists := xlFile.Sheet[options.SheetName]; !exists {
			return sheet, fmt.Errorf("no sheet named \"%s\" available", options.SheetName)
		}

		sheet = xlFile.Sheet[options.SheetName]
	} else {
		// Sheets length
		numSheets := len(xlFile.Sheets)

		switch {
		case numSheets == 0:
			return sheet, fmt.Errorf("this XLSX file contains no sheets")
		case options.SheetIndex >= numSheets:
			return sheet, fmt.Errorf("no sheet %d available, please select a sheet between 0 and %d", options.SheetIndex, numSheets-1)
		}

		sheet = xlFile.Sheets[options.SheetIndex]
	}

	return sheet, nil
}

func setRows(sheet *xlsx.Sheet) ([]*xlsx.Row, error) {
	var rows []*xlsx.Row
	if options.StartLine > -1 {
		// Read partial sheet
		numRows := len(sheet.Rows)
		switch {
		case numRows == 0:
			return rows, fmt.Errorf("this worksheet contains no rows")
		case options.StartLine >= numRows:
			return rows, fmt.Errorf("no row %d available, please select a row between 0 and %d", options.HeaderLine, numRows-1)
		}

		if options.Limit > 0 {
			options.EndLine = options.StartLine + options.Limit
		}

		if options.Limit == 0 || options.EndLine > numRows {
			options.EndLine = numRows
		}

		rows = sheet.Rows[options.StartLine:options.EndLine]
		if options.HeaderLine > -1 {
			if options.HeaderLine >= numRows {
				return rows, fmt.Errorf("no row %d available, please select a row between 0 and %d", options.HeaderLine, numRows-1)
			}

			rows = append(sheet.Rows[options.HeaderLine:options.HeaderLine+1], rows...)
		}
	} else {
		rows = sheet.Rows
	}

	return rows, nil
}

func outputCsv(csvWriter *csv.Writer, rows []*xlsx.Row) {
	// Read sheet
	for _, row := range rows {
		var record []string
		if row != nil {
			for _, cell := range row.Cells {
				record = append(record, cell.String())
			}

			csvWriter.Write(record)
		}
	}
}
