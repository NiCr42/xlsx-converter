# XSLX Converter

Simple file converter : Print a sheet from a XLSX file to standard ouput (Stdout) as CSV format.

## Installation

Use Go dependency management tool https://golang.github.io/dep/

```shell
go get github.com/nicr42/xlsx-converter
dep ensure -update
```

## Usage

```shell
Usage: [-output-file <file> (-sheet-index <index> | -sheet-name <name>) -header-line <line> -start-line <line> -limit <limi
t-number-of-lines>] <file>
  -header-line int
        Index of header line. (default -1)
  -limit int
        Limit number of lines to retrieve.
  -output-file string
        CSV output file.
  -sheet-index int
        Index of worksheet.
  -sheet-name string
        Name of worksheet.
  -start-line int
        Index of start line. (default -1)

```
