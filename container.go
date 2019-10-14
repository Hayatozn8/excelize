package excelize

import (
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
)

//xlContainer{xlBook{xlSheet}}

type xlContainer struct {
	bk_pool map[string]*xlBook
}

type xlBook struct {
	bkApp           *File
	sheetCollection map[string]*xlSheet
}

type xlSheet struct {
	sheetApp     *xlsxWorksheet
	activate_row int
	activate_col int
}

var container *xlContainer
var once sync.Once

//////////////////////container
func GetXlContainer() *xlContainer {
	once.Do(func() {
		container = &xlContainer{
			bk_pool: make(map[string]*xlBook),
		}
	})

	return container
}

func (xlc *xlContainer) MakeBook(path string, sheetNames ...string) (*xlBook, error) {
	// change to abs path
	path, err := filepath.Abs(path)
	if err != nil {
		return nil, err
	}

	// check container book pool
	if _, has := container.bk_pool[path]; has {
		return nil, errors.New("this file has been loaded or made, please use method: LoadBook()")
	}

	// check file exist
	_, err = os.Stat(path)
	if err == nil || os.IsExist(err) {
		return nil, errors.New("this file already exists, please use method: LoadBook()")
	}

	// make a new book
	f := NewFile()

	// make sheet
	if len(sheetNames) != 0 {
		f.SetSheetName("Sheet1", sheetNames[0])
		for _, sheetName := range sheetNames[1:] {
			f.NewSheet(sheetName)
		}
	}

	// save file
	err = f.SaveAs(path)
	if err != nil {
		fmt.Println(err)
	}

	// load file
	f, err = OpenFile(path)
	if err != nil {
		return nil, err
	}

	// new xlBook and return
	xlbk := &xlBook{
		bkApp:           f,
		sheetCollection: make(map[string]*xlSheet),
	}
	container.bk_pool[path] = xlbk

	return xlbk, nil
}

func (xlc *xlContainer) LoadBook(path string) (*xlBook, error) {
	// change to abs path
	path, err := filepath.Abs(path)
	if err != nil {
		return nil, err
	}

	// check container book pool
	if f, has := container.bk_pool[path]; has {
		return f, nil
	}

	// check file exist
	_, err = os.Stat(path)
	if err != nil && os.IsNotExist(err) {
		return nil, errors.New("this file is not exist, please use method: MakeBook()")
	}

	// open file
	f, err := OpenFile(path)
	if err != nil {
		return nil, err
	}

	// new xlBook and return
	xlbk := &xlBook{
		bkApp:           f,
		sheetCollection: make(map[string]*xlSheet),
	}
	container.bk_pool[path] = xlbk

	return xlbk, nil
}

func (xlc *xlContainer) CopyBook(from, to string) error {
	// change to abs path
	from, err := filepath.Abs(from)
	if err != nil {
		return err
	}

	to, err = filepath.Abs(to)
	if err != nil {
		return err
	}

	// from file is used????
	//if _, has := xlc.bk_pool[]; has{
	// TODO
	//}

	// open from file
	fromFile, err := os.Open(from)
	if err != nil {
		return err
	}
	defer fromFile.Close()

	// open to file
	toFile, err := os.Create(to)
	if err != nil {
		return err
	}
	defer toFile.Close()

	// copy
	_, err = io.Copy(toFile, fromFile)

	return err
}

func (xlc *xlContainer) PoolCount() int {
	return len(xlc.bk_pool)
}

func (xlc *xlContainer) ClearAllBk() {
	if len(xlc.bk_pool) != 0 {
		xlc.bk_pool = make(map[string]*xlBook)
	}
}

func (xlc *xlContainer) ClearBook(path string) error {
	// lock every book (for every book, check: exist other lock)
	if len(xlc.bk_pool) == 0 {
		return nil
	}
	// change to abs path
	path, err := filepath.Abs(path)
	if err != nil {
		return err
	}
	// clear book
	delete(xlc.bk_pool, path)
	return nil
}

///////////////////////////////////////////////////////////////////////////////// book
// index of sheet from 1???
func (bk *xlBook) CopySheetByName(from, to string) error {
	// check from sheetName
	fromIndex := bk.bkApp.GetSheetIndex(trimSheetName(from))
	if fromIndex == 0 {
		return errors.New("can not find from sheet")
	}

	// get toIndex
	toIndex := bk.bkApp.NewSheet(trimSheetName(to))

	// copy
	return bk.bkApp.copySheet(fromIndex, toIndex)
}

func (bk *xlBook) NewSheet(sheetName string) (*xlSheet, error) {
	bk.bkApp.NewSheet(sheetName)

	// save in sheetCollection
	return bk.initSheetApp(sheetName)
}

func (bk *xlBook) Save() error {
	return bk.bkApp.Save()
}

func (bk *xlBook) SaveAs(path string) error {
	return bk.bkApp.SaveAs(path)
}

//sheetName <--> sheetIndex
func (bk *xlBook) GetSheet(sheetName string) (*xlSheet, error) {
	// get sheet from sheetCollection
	sheetName = trimSheetName(sheetName)
	if xls, has := bk.sheetCollection[sheetName]; has{
		return xls, nil
	}

	// save in sheetCollection
	return bk.initSheetApp(sheetName)
}


func (bk *xlBook) initSheetApp(trimedSheetName string) (*xlSheet, error) {
	xlsApp, err := bk.bkApp.workSheetReader(trimedSheetName)
	if err != nil {
		return nil, err
	}

	xls := &xlSheet{
		sheetApp:     xlsApp,
		activate_row: 1,
		activate_col: 1,
	}

	bk.sheetCollection[trimedSheetName] = xls

	return xls, nil
}
///////////////////////////////////////////////////////////////////////////////// sheet
/*
container.bk
defer bk.save()


primitive
bk copysheet

sheet.do
sheet-set
sheet.Do("set", "img", "startrange", "val")
sheet.Do("set", "val", "range", "val")

sheet-get
sheet.Do("get", "val", "range", "val")
*/

//type xlSheet struct {
//	sheetApp *xlsxWorksheet
//	activate_row int
//	activate_col int
//}

func (xls *xlSheet) mergeCellsParser(axis string) (string, error) {
	axis = strings.ToUpper(axis)
	if xls.sheetApp.MergeCells != nil {
		for i := 0; i < len(xls.sheetApp.MergeCells.Cells); i++ {
			ok, err := checkCellInArea(axis, xls.sheetApp.MergeCells.Cells[i].Ref)
			if err != nil {
				return axis, err
			}
			if ok {
				axis = strings.Split(xls.sheetApp.MergeCells.Cells[i].Ref, ":")[0]
			}
		}
	}
	return axis, nil
}

func (xls *xlSheet) prepareCellStyle(col, style int) int {
	if xls.sheetApp.Cols != nil && style == 0 {
		for _, c := range xls.sheetApp.Cols.Col {
			if c.Min <= col && col <= c.Max {
				style = c.Style
			}
		}
	}
	return style
}

//interface{}
func (xlsht *xlSheet) Set(target string, xlrange string, value int) error {
	// change range to Coord(slice)
	coord, err := RangeToCoord(xlrange)
	if err != nil {
		return err
	}

	// // find method
	// m := reflect.ValueOf(xlsht).Elem().MethodByName("set" + strings.ToLower(target))
	// if m.IsNil() {
	// 	return errors.New("can not find method, please check target")
	// }

	// // make paramters
	// params := []reflect.Value{
	// 	reflect.ValueOf(coord),
	// 	reflect.ValueOf(value),
	// }

	// // run method
	// res := m.Call(params)
	// return res[0].Interface().(error)

	xlrange, err = xlsht.mergeCellsParser(xlrange)
	if err != nil {
		return err
	}
	col, row := coord[0], coord[1]

	prepareSheetXML(xlsht.sheetApp, col, row)

	cellData := &xlsht.sheetApp.SheetData.Row[row-1].C[col-1]

	//
	cellData.S = xlsht.prepareCellStyle(col, cellData.S)
	cellData.T = ""
	cellData.V = strconv.Itoa(value)
	return err
}

// func (xlsht *xlSheet) setval(coord []int, value interface{}) (error){
// 	left, up := coord[0], coord[1]
// 	for

// }

// func (xlsht *xlSheet) Get(target string, xlrange string)(interface{}, error){

// }

func RangeToCoord(xlrange string) ([]int, error) {
	res := strings.Split(xlrange, ":")

	// TODO repalce space?????

	if count := len(res); count == 1 {
		// type: "a2"
		col, row, err := CellNameToCoordinates(res[0])
		if err != nil {
			return nil, err
		}

		return []int{col, row}, nil
	} else if count == 2 {
		//type : "a1:b2"
		left, up, err := CellNameToCoordinates(res[0])
		if err != nil {
			return nil, err
		}

		right, down, err := CellNameToCoordinates(res[1])
		if err != nil {
			return nil, err
		}

		return []int{left, up, right, down}, nil
	} else {
		// other is errror
		return nil, errors.New("can not analyse range")
	}
}

