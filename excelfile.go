package excelworker

import "strings"

type ExcelFile struct {
	path string
	data [][]string
}

func NewExcelFile(filePath string, data [][]string) *ExcelFile {
	return &ExcelFile{
		path: filePath,
		data: data,
	}
}

func (ef *ExcelFile) ChangeFormat(ToFormat string) {
	splitted := strings.Split(ef.path, ".")
	splitted[len(splitted)-1] = ToFormat
	ef.path = strings.Join(splitted, ".")
}

func (ef ExcelFile) GetData() [][]string {
	return ef.data
}

func (ef *ExcelFile) GetFileName() string {
	splitted := strings.Split(ef.GetPath(), "/")
	return splitted[len(splitted)-1]
}

func (ef ExcelFile) GetPath() string {
	return ef.path
}
