// Copyright 2016 - 2019 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.
//
// Package excelize providing a set of functions that allow you to write to
// and read from XLSX files. Support reads and writes XLSX file generated by
// Microsoft Excel™ 2007 and later. Support save file without losing original
// charts of XLSX. This library needs Go version 1.10 or later.

package excelize

import (
	"reflect"
	"errors"
	_"fmt"
)

func (f *File) SetRangeValue(sheet, axis string, values interface{}) error{
	rangeVal := reflect.ValueOf(values)

	if rangeVal.Kind() != reflect.Slice {
		return errors.New("pointer to slice expected")
	}

	xlsx, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}

	startCol, startRow, err := CellNameToCoordinates(axis)
	if err != nil {
		return err
	}

	// a := [][]string{{"a","b","c"}, {"d", "r", "y", "t"}, {"c", "x"}, {"u"}}
	// a := []interface{}{[]string{"a","b","c"}, 2, 3, []string{"s","x"}}
	for i := 0; i < rangeVal.Len(); i++{
		rowVal := rangeVal.Index(i)

		// a := []interface{}{[]string{"a","b","c"}, 2, 3, []string{"s","x"}}
		rowValKind := rowVal.Kind()
		if rowValKind == reflect.Interface {
			rowVal = reflect.ValueOf(rowVal.Interface())
			rowValKind = rowVal.Kind()
		}

		if rowValKind != reflect.Slice{
			cell, err := CoordinatesToCellName(startCol, startRow+i)
			if err != nil {
				return err
			}

			cellData, col, row, err := f.prepareCell(xlsx, cell)
			if err != nil {
				return err
			}

			if err := f.setCellValue(xlsx, cellData, col, row, rowVal.Interface()); err != nil {
				return err
			}
			// if err := f.SetCellValue(sheet, cell, rowVal.Interface()); err != nil {
			// 	return err
			// }
			continue
		}
		
		for j := 0; j < rowVal.Len(); j++ {
			cell, err := CoordinatesToCellName(startCol+j, startRow+i)
			if err != nil {
				return err
			}

			cellData, col, row, err := f.prepareCell(xlsx, cell)
			if err != nil {
				return err
			}

			if err := f.setCellValue(xlsx, cellData, col, row, rowVal.Index(j).Interface()); err != nil {
				return err
			}

			// if err := f.SetCellValue(sheet, cell, rowVal.Index(j).Interface()); err != nil {
			// 	return err
			// }
		}
	}
	return nil

}
