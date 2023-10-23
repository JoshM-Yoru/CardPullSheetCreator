package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {

	inputFlag := flag.String("i", "", "Path to the input file")
	outputFlag := flag.String("o", "", "Name of output file")

	flag.Parse()

	if *inputFlag == "i" {
		fmt.Println("Please provide a file using the -i flag.")
		os.Exit(1)
	}
	if *outputFlag == "o" {
		fmt.Println("Please provide a file using the -o flag.")
		os.Exit(1)
	}

	file, err := os.Open(*inputFlag)
	if err != nil {
		fmt.Println("Error opening this file.")
		os.Exit(0)
	}

	outputFile := excelize.NewFile()

	defer file.Close()

	reader := csv.NewReader(file)

	reader.Comma = ','

	records, err := reader.ReadAll()
	if err != nil {
		fmt.Println("Error reading the file: ", err)
		os.Exit(1)
	}

	for _, record := range records {

		for _, field := range record {
			cleanedLine := strings.Replace(field, "\"", "", -1)
			fmt.Println(cleanedLine)
		}
	}

	matrixQuickSort(records, 1, len(records)-2)

	excelWriter(outputFile, records)

	outputFileName := *outputFlag + ".xlsx"

	if err := outputFile.SaveAs(outputFileName); err != nil {
		fmt.Println("Unable to save file: ", err)
		return
	} else {
		fmt.Printf("Excel file created: %s.xlsx", *outputFlag)
	}

}

func excelWriter(file *excelize.File, matrix [][]string) {
	for i := 0; i < len(matrix); i++ {

		for j := 0; j < len(matrix[0]); j++ {
			file.SetCellStr("Sheet1", string(rune('A'+j))+strconv.Itoa(i+1), matrix[i][j])
		}
	}
}

func matrixQuickSort(matrix [][]string, low int, high int) {

	if low < high {
		pivot := partition(matrix, low, high)

		matrixQuickSort(matrix, low, pivot-1)
		matrixQuickSort(matrix, pivot+1, high)
	}

}

func partition(matrix [][]string, low, high int) int {

	pivot := matrix[high][0]

	i := low - 1

	for j := low; j <= high-1; j++ {

		if matrix[j][0] <= pivot {
			i++
			matrix[i][0], matrix[j][0] = matrix[j][0], matrix[i][0]
		}

	}

	matrix[i+1], matrix[high] = matrix[high], matrix[i+1]

	return i + 1
}
