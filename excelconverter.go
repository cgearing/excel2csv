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

func writeXLasCSV(filename string, ignoreHidden bool) error {
	file, err := excelize.OpenFile(filename)
	if err != nil {
		return err
	}
	sheets := file.GetSheetMap()
	fileNameWithExtension := path.Base(filename)
	extension := path.Ext(filename)
	fileNameNoExt := fileNameWithExtension[:len(fileNameWithExtension)-len(extension)]
	for index, sheetName := range sheets {
		visible := file.GetSheetVisible(sheetName)
		if visible == false && ignoreHidden == true {
			return nil
		}
		newFileName := fmt.Sprintf("output/%0s-%1s.csv", fileNameNoExt, sheetName)
		newFile, err := os.OpenFile(newFileName, os.O_RDWR|os.O_CREATE, 0755)
		if err != nil {
			fmt.Println("can't create new file", newFileName)
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

func walkFunc(path string, info os.FileInfo, err error, ignoreHidden bool) error {
	if err != nil {
		return err
	}
	if !info.IsDir() && filepath.Ext(info.Name()) == ".xlsx" {
		err := writeXLasCSV(path, ignoreHidden)
		if err != nil {
			return err
		}
	}
	return nil
}

func wrapWalkFunc(ignoreHidden bool) filepath.WalkFunc {
	return func(path string, info os.FileInfo, err error) error {
		return walkFunc(path, info, err, ignoreHidden)
	}
}

func parseInputDir() (string, bool) {
	inputDirFlag := flag.String("i", "input", "input directory of xlsx files")
	ignoreHidden := flag.Bool("h", true, "ignore hidden xlsx sheets")
	flag.Parse()
	return *inputDirFlag, *ignoreHidden
}

func ensureOutputDir() error {
	outputDir := "./output"
	return os.Mkdir(outputDir, 0777)
}

func main() {
	inputDir, ignoreHidden := parseInputDir()

	ensureOutputDir()

	wrappedWalkFunc := wrapWalkFunc(ignoreHidden)

	err := filepath.Walk(inputDir, wrappedWalkFunc)
	if err != nil {
		log.Fatal("Fatal error", err)
	}
}
