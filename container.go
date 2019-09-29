package excelize

import (
	"errors"
	"fmt"
	"os"
	"sync"
)

type xlContainer struct {
	bk_pool map[string]*File
}

var container *xlContainer
var once sync.Once

func GetXlContainer() *xlContainer {
	once.Do(func() {
		container = &xlContainer{
			bk_pool: make(map[string]*File),
		}
	})

	return container
}

func (xlc *xlContainer) MakeBook(path string, sheetNames ...string) (*File, error) {
	// check container book pool
	if _, has := container.bk_pool[path]; has {
		return nil, errors.New("this file has been loaded or made, please use method: LoadBook()")
	}

	// check file exist
	if _, err := os.Stat(path); err == nil || os.IsExist(err) {
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
	err := f.SaveAs(path)
	if err != nil {
		fmt.Println(err)
	}

	// load file
	f, err = OpenFile(path)
	if err != nil {
		return nil, err
	}

	container.bk_pool[path] = f
	return f, nil
}

func (xlc *xlContainer) LoadBook(path string) (*File, error) {
	// check container book pool
	if f, has := container.bk_pool[path]; has {
		return f, nil
	}

	// check file exist
	if _, err := os.Stat(path); err != nil && os.IsNotExist(err) {
		return nil, errors.New("this file is not exist, please use method: MakeBook()")
	}

	// open file
	f, err := OpenFile(path)
	if err != nil {
		return nil, err
	}

	container.bk_pool[path] = f
	return f, nil
}

func (bk *File) GetSheet(sheetName string) (*xlsxWorksheet, error) {
	return bk.workSheetReader(sheetName)
}

//XlBook
//

/*
XlSheet
set：
	value
	image
	formula
get：
	value

activate_cell
used_range
used_range_address_str
used_range_address_inttuple
used_range_value
/*
