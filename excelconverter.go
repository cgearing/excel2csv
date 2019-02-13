package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"log"
	"os"
	"path"
	"path/filepath"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func writeXLasCSV(filename string) error {
	file, err := excelize.OpenFile(filename)
	if err != nil {
		return err
	}
	sheets := file.GetSheetMap()
	fileNameWithExtension := path.Base(filename)
	extension := path.Ext(filename)
	fileNameNoExt := fileNameWithExtension[: len(fileNameWithExtension)-len(extension)]

	for index, sheetName := range sheets {
		newFileName := fmt.Sprintf("output/%0s-%1s.csv", fileNameNoExt, sheetName)
		newFile, err := os.OpenFile(newFileName, os.O_RDWR|os.O_CREATE, 0755)
		if err != nil {
			return err
		}
		csvWriter := csv.NewWriter(newFile)
		fmt.Println(index, sheetName)
		err = csvWriter.WriteAll(file.GetRows(sheetName))
		if err != nil {
			return err
		}
	}
	return nil
}

func walkFunc(path string, info os.FileInfo, err error) error {
	if err != nil {
		return err
	}
	if !info.IsDir() && filepath.Ext(info.Name()) == ".xlsx" {
		err := writeXLasCSV(path)
		if err != nil {
			return err
		}
	}
	return nil
}

func parseInputDir() string {
	inputDirFlag := flag.String("i", "input", "input directory of xlsx files")
	flag.Parse()
	return *inputDirFlag
}

func ensureOutputDir() error {
	outputDir := "./output"
	return os.Mkdir(outputDir, 0777)
}

func main() {
	inputDir := parseInputDir()
	ensureOutputDir()

	err := filepath.Walk(inputDir, walkFunc)
	if err != nil {
		log.Fatal("Fatal error", err)
	}
}
