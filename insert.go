package main

import (
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main(){
	xlsF := `./goMysql.xlsx`
	sheetName := "equip.main"
	nf, err := os.Create("insert.txt")
	if err != nil{
		panic(err)
	}
	f, err := excelize.OpenFile(xlsF)
	if err != nil{
		panic(err)
	}
	rows := f.GetRows(sheetName)
	var syntax string
	var titles string
	for i, row:= range rows{
		if i == 0 {
			for _, col := range row{
				titles += fmt.Sprint(``+ col +`,`)
			}
		}
		if i > 0{
			var values string
			for _, col := range row{
				v := "null"
				if col != ""{
					v = fmt.Sprintf("'%s'",col)
				}
				values += fmt.Sprintf("%s,", v)
			}
			syntax += fmt.Sprintf(
				"INSERT INTO %s(%s)VALUES(%s); \n",
				sheetName,strings.TrimSuffix(titles,","), strings.TrimSuffix(values,","))
		}
	}
	defer nf.Close()
	io.Copy(nf,strings.NewReader(syntax))

}
