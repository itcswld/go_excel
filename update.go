package main

import (
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main(){

	xlsF := `D:\PH\Equip\goMysql.xlsx`
	sheetName := "equip.main"
	nf, err := os.Create("update.txt")
	if err != nil {
		fmt.Println(err.Error())
	}
	f, err := excelize.OpenFile(xlsF)
	if err != nil{
		panic(err)
	}
	rows := f.GetRows(sheetName)
	var col_title []string
	var syntax string

	for i, row := range rows {

		if i == 0 {
			for _, col := range row {
				col_title = append(col_title, col)
			}
		}
		if i > 0{
			//upate
			r := row[1:len(row)]
			for i = 0;  i < len(r); i++ {
				v := "null"
				if  r[i] != "" {
					v = r[i]
				}
				syntax += fmt.Sprint(`
				UPDATE equip.equip SET `+ col_title[i+1] +`= '`+ v +`'
				WHERE `+ col_title[0] +`= '`+ row[0] +`';
				`)
			}
		}
	}
	defer nf.Close()
	io.Copy(nf,strings.NewReader(syntax))
}
