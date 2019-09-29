package excelize

import (
	"testing"
)

func TestLoadBook(t *testing.T) {
	a := GetXlContainer()
	p := "/Users/liujinsuo/gosys/others/yy.xlsx"
	f, err := a.LoadBook(p)
	if err != nil {
		t.Fatalf("err=%v", err)
	}
	// b := f.GetSheetMap()
	// fmt.Println(b)
	f.SetCellValue("123", "B2", 250)
	f.Save()
}

func TestMakeBook(t *testing.T) {
	a := GetXlContainer()
	p := "./test/yy.xlsx"
	f, err := a.MakeBook(p, "123")
	if err != nil {
		t.Fatalf("err=%v", err)
	}
	// b := f.GetSheetMap()
	// fmt.Println(b)
	// f.SetCellValue("123", "B2", 100)
	// f.SaveAs("/Users/liujinsuo/gosys/others/zzz.xlsx")
	f.NewSheet("999")
	f.Save()
}
