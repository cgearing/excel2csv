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
	fileNameNoExt := fileNameWithExtension[0 : len(fileNameWithExtension)-len(extension)]
	for _, sheetName := range sheets {
		fnames := []interface{}{fileNameNoExt, sheetName}
		newFileName := fmt.Sprintf("output/%0s-%1s.csv", fnames[0], fnames[1])
		newFile, err := os.OpenFile(newFileName, os.O_RDWR|os.O_CREATE, 0755)
		if err != nil {
			return err
		}
		csvWriter := csv.NewWriter(newFile)
		for index, name := range sheets {
			fmt.Println(index, name)
			err := csvWriter.WriteAll(file.GetRows(name))
			if err != nil {
				return err
			}
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

func main() {
	inputDir := flag.String("i", "input", "input directory of xlsx files")
	flag.Parse()
	rootPath := *inputDir
	err := filepath.Walk(rootPath, walkFunc)
	if err != nil {
		log.Fatal("Fatal error", err)
	}
}
