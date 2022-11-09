package excelworker

import (
	"errors"
	"fmt"
	"io/fs"
	"io/ioutil"
	"os"
	"strings"

	"github.com/extrame/xls"
	"github.com/xuri/excelize/v2"
)

type ExcelWorker interface {
	ReadExcel(fileName string) (*ExcelFile, error)
	GetColumn(data [][]string, column int) []string
	WriteExcel(ef *ExcelFile) error
	SaveByteAsFile(file []byte, fileName string, force bool) error
	// columns starts from zero
	UpdateExcelColumn(ef *ExcelFile, column int, fin func(string) string) error
	UpdateExcelRows(ef *ExcelFile, fin func([]string) []string) error
}

type excelWorker struct {
}

func NewXlsWorker() ExcelWorker {
	return &excelWorker{}
}

func (x excelWorker) UpdateExcelColumn(ef *ExcelFile, column int, fin func(string) string) error {
	res := ef.data
	for i := range res {
		if len(res[i]) >= column {
			res[i][column] = fin(res[i][column])
		}
	}
	return nil
}
func (x excelWorker) UpdateExcelRows(ef *ExcelFile, fin func([]string) []string) error {
	res := ef.data
	for i := range res {
		res[i] = fin(res[i])
	}
	return nil
}

func (x excelWorker) WriteExcel(ef *ExcelFile) error {
	x.WriteToXlsx(ef)
	return nil
}

func (x excelWorker) WriteToXlsx(ef *ExcelFile) error {

	f := excelize.NewFile()
	sheetName := "Sheet1"
	index := f.NewSheet(sheetName)
	f.SetActiveSheet(index)
	data := ef.data
	for i := 0; i < len(data); i++ {
		for j := 0; j < len(data[i]); j++ {
			row := fmt.Sprintf("%v%v", x.StringOfColumn(j), i+1)
			f.SetCellValue(sheetName, row, data[i][j])
		}
	}
	if err := f.SaveAs(ef.path); err != nil {
		return err
	}
	return nil
}

func (x excelWorker) ReadExcel(fileName string) (*ExcelFile, error) {
	splitted := strings.Split(fileName, ".")
	format := splitted[len(splitted)-1]
	ef := &ExcelFile{path: fileName}
	if format == "xlsx" {
		data, err := x.ReadXlsx(fileName)
		ef.data = data
		return ef, err
	} else if format == "xls" {
		data, err := x.ReadXls(fileName)
		ef.data = data
		return ef, err
	}
	return ef, errors.New("file is not .xls or .xlsx")
}

func (x excelWorker) ReadXls(fileName string) ([][]string, error) {
	fmt.Println(fileName)
	f, e := xls.Open(fileName, "utf-8")
	if e != nil {
		return [][]string{}, fmt.Errorf("cannot open file %s", fileName)
	}
	res := f.ReadAllCells(10000000000000)
	return res, nil
}

func (x excelWorker) ReadXlsx(fileName string) ([][]string, error) {
	f, err := excelize.OpenFile(fileName, excelize.Options{})
	if err != nil {
		f.Close()
		return [][]string{}, err
	}
	defer f.Close()
	if err != nil {
		return [][]string{}, err
	}
	sheetName := f.GetSheetName(0)
	res, err := f.GetRows(sheetName)
	if err != nil {
		return [][]string{}, err
	}
	return res, nil
}

func (x excelWorker) SaveXlsAsXlsx(fileName string) error {
	data, err := x.ReadXls(fileName)
	if err != nil {
		return err
	}
	x.WriteToXlsx(&ExcelFile{path: fileName, data: data})
	if err != nil {
		return err
	}
	return nil
}

func (ew excelWorker) GetColumn(data [][]string, column int) []string {
	out := make([]string, 0)
	for i := 0; i < len(data); i++ {
		out = append(out, data[i][column])
	}
	return out
}

func (excelWorker) StringOfColumn(column int) string {
	return string(column + 65)
}

func (x excelWorker) SaveByteAsFile(file []byte, fileName string, force bool) error {
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
	_, err = x.ReadExcel(fileName)
	if err != nil {
		err := os.Remove(fileName)
		if err != nil {
			return err
		}
		return fmt.Errorf("cannot create file with this format, %s", err)
	}
	return nil
}
