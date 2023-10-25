package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {

    fileName := os.Args[1]

    fileName = filepath.ToSlash(fileName)

	file, err := os.Open(fileName)
	if err != nil {
		fmt.Println("Error opening this file.")
		os.Exit(0)
	}

    fmt.Println(fileName)

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
			strings.Replace(field, "\"", "", -1)
		}
	}

	matrixQuickSort(records, 1, len(records)-2, 0)

	for i := 1; i < len(records)-1; i++ {

		low := 1

		if i == len(records)-2 {
			matrixQuickSort(records, low, len(records)-2, 1)
		}

		if records[i][0] != records[i-1][0] && i-low > 1 {
			matrixQuickSort(records, low, i, 1)
		}
	}

	excelWriter(outputFile, records)

    currentTime := time.Now()

    formattedTime := currentTime.Format(time.RFC822)

    output := "PullSheet_" + strings.Replace(formattedTime, " ", "_", -1)

	outputFileName := output + ".xlsx"

	if err := outputFile.SaveAs(outputFileName); err != nil {
		fmt.Println("Unable to save file: ", err)
		return
	} else {
		fmt.Println("Excel file created: ", outputFileName)
	}

}

func excelWriter(file *excelize.File, matrix [][]string) {
	for i := 0; i < len(matrix); i++ {

		for j := 0; j < len(matrix[0]); j++ {
			file.SetCellStr("Sheet1", string(rune('A'+j))+strconv.Itoa(i+1), matrix[i][j])
		}
	}
}

func matrixQuickSort(matrix [][]string, low int, high int, col int) {

	if low < high {
		pivot := partition(matrix, low, high, col)

		matrixQuickSort(matrix, low, pivot-1, col)
		matrixQuickSort(matrix, pivot+1, high, col)
	}

}

func partition(matrix [][]string, low int, high int, col int) int {

	pivot := matrix[high][col]

	i := low - 1

	for j := low; j <= high-1; j++ {

		if matrix[j][col] <= pivot {
			i++
			matrix[i][col], matrix[j][col] = matrix[j][col], matrix[i][col]
		}

	}

	matrix[i+1], matrix[high] = matrix[high], matrix[i+1]

	return i + 1
}

// func groupByGame(matrix [][]string) {
//
//     var games map[string]int
//
//     for i := 0; i < len(matrix); i++ {
//         if _, exists := games[matrix[i][0]]; !exists {
//             games[matrix[i][0]] = 1
//         } else {
//             games[matrix[i][0]] += 1
//         }
//     }
// }
