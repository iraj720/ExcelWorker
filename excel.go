package excelworker

import (
	"errors"
	"fmt"
	"io/fs"
	"io/ioutil"
	"os"

	"github.com/extrame/xls"
	"github.com/xuri/excelize/v2"
)

type xlsWorker struct {
}

type XlsWorker interface {

	// columns starts from zero
	UpdateXlsColumn(fileName string, column int, fin func(string) string) error
}

func NewXlsWorker() XlsWorker {
	return &xlsWorker{}
}

func (x xlsWorker) UpdateXlsColumn(fileName string, column int, fin func(string) string) error {
	res, err := x.ReadXls(fileName)
	if err != nil {
		return err
	}
	err = x.WriteToXlsx(fmt.Sprintf("out_%v", fileName), res)
	if err != nil {
		return err
	}
	return nil
}

func (x xlsWorker) WriteToXlsx(fileName string, data [][]string) error {
	f := excelize.NewFile()
	sheetName := "Sheet1"
	index := f.NewSheet(sheetName)
	f.SetActiveSheet(index)
	for i := 0; i < len(data); i++ {
		for j := 0; j < len(data[i]); j++ {
			row := fmt.Sprintf("%v%v", x.StringOfColumn(j), i+1)
			f.SetCellValue(sheetName, row, data[i][j])
		}
	}
	if err := f.SaveAs(fileName + ".xlsx"); err != nil {
		return err
	}
	return nil
}

func (x xlsWorker) ReadXls(fileName string) ([][]string, error) {
	f, e := xls.Open(fileName, "utf-8")
	if e != nil {
		return [][]string{}, fmt.Errorf("cannot open file %s", fileName)
	}
	res := f.ReadAllCells(1000000000)
	return res, nil
}

func (x xlsWorker) ReadXlsx(fileName string) ([][]string, error) {
	f, err := excelize.OpenFile(fileName, excelize.Options{})
	if err != nil {
		return [][]string{}, err
	}
	res, err := f.GetRows(fileName)
	if err != nil {
		return [][]string{}, err
	}
	return res, nil
}

func (x xlsWorker) SaveXlsAsXlsx(fileName string) error {
	data, err := x.ReadXls(fileName)
	if err != nil {
		return err
	}
	x.WriteToXlsx(fileName+".xlsx", data)
	if err != nil {
		return err
	}
	return nil
}

func (x xlsWorker) SaveByteAsFile(file []byte, fileName string, force bool) error {
	err := ioutil.WriteFile(fileName, file, fs.ModeAppend)
	if err != nil {
		_, err = os.Stat(fileName)
		if !errors.Is(err, os.ErrNotExist) && force {
			err := os.Remove(fileName)
			if err != nil {
				return err
			}
		} else {
			return err
		}
	}
	return nil
}

func (xlsWorker) StringOfColumn(column int) string {
	return string(column + 65)
}
