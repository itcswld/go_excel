package main

import (
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var (
	// xlsF = `D:\PM\Project\2023_LC_WMS專案\DB\ITWMSDB\old_new.xlsx`
	// xlsF      = `D:\PM\Project\2023_LC_WMS專案\測試\itwmsdb\IT_0922.xlsx`
	xlsF      = `D:\PM\Project\2023_LC_WMS專案\上線\0927\lc_wms保存天數.xlsx`
	sheetName = "moci_pro_expd"
	// sheetName = "OTRT"
)

func main() {
	fn := fmt.Sprintf("%s.txt", sheetName)
	nf, err := os.Create(fn)
	if err != nil {
		panic(err)
	}
	f, err := excelize.OpenFile(xlsF)
	if err != nil {
		panic(err)
	}
	rows := f.GetRows(sheetName)
	var syntax string
	var columname string
	for i, row := range rows {
		if i == 0 {
			for _, col := range row {
				columname += fmt.Sprint(`` + col + `,`)
			}
		}
		if i > 0 {
			var values string
			for _, col := range row {
				v := "''"
				if col != "" {
					if len(col) > 2 && col[len(col)-2:] == "()" || len(col) > 2 && col[:1] == "@" || col == "NULL" {
						v = fmt.Sprintf("%s", col)
					} else {
						v = fmt.Sprintf("'%s'", col)
					}
				}
				values += fmt.Sprintf("%s,", v)
			}
			syntax += fmt.Sprintf(
				"INSERT INTO %s(%s)VALUES(%s); \n",
				sheetName, strings.TrimSuffix(columname, ","), strings.TrimSuffix(values, ","))
		}
	}
	defer nf.Close()
	io.Copy(nf, strings.NewReader(syntax))

}
