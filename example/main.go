package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"os"
	"time"

	"strings"

	"github.com/regorov/rbuilder"
	"github.com/tealeg/xlsx"
)

func main() {

	if len(os.Args) != 4 {
		fmt.Println("rbuilder temlpate.xlsx data.json output.xlsx")
		os.Exit(1)
	}

	tmpl, err := xlsx.OpenFile(os.Args[1])
	if err != nil {
		fmt.Println(err)
		os.Exit(2)
	}

	buf, err := ioutil.ReadFile(os.Args[2])
	if err != nil {
		fmt.Println(err)
		os.Exit(3)
	}

	s := map[string]interface{}{"License": "Лицензия 123-456-78 от 27.09.13 г.",
		"CompanyName": "ООО \"Элефант софт\" г.Казань",
		"CurrentTime": time.Now().Format("02.01.2006 15:04:05")}

	rbt := rbuilder.NewTemplate(tmpl, s)

	fmt.Println(string(buf))

	var a map[string]interface{}

	err = json.Unmarshal(buf, &a)
	if err != nil {
		fmt.Printf("json unmarshal: %+v", err)
		os.Exit(3)
	}

	fmt.Println("Start rendering")

	result, err := rbt.Render(a)
	if err != nil {
		fmt.Println(err)
		os.Exit(3)
	}

	mergeCells(result)

	err = result.Save(os.Args[3])
	if err != nil {
		fmt.Println(err)
		os.Exit(3)
	}

}

func mergeCells(report *xlsx.File) {

	br, bc := 0, 0
	b := false
	for s := range report.Sheets {
		for r := range report.Sheets[s].Rows {
			for c := range report.Sheets[s].Rows[r].Cells {
				val := report.Sheets[s].Rows[r].Cells[c].Value

				if b == false {
					if strings.Index(val, "<<") != 0 {
						continue
					}
					br, bc, b = r, c, true
					report.Sheets[s].Rows[r].Cells[c].Value = val[2:]

					continue
				}

				// сюда попадает только если найдено начало требования
				// по объединению ячеек

				if !strings.Contains(val, ">>") {
					continue
				}
				println("<< found", s, r, c)

				// сюда попадает, если найдено требование окончания
				// сегмента по объединению ячеек

				//if left {
				report.Sheets[s].Rows[br].Cells[bc].Merge(c-bc, r-br)

				/*} else {
				val := report.Sheets[s].Rows[r].Cells[c].Value

				report.Sheets[s].Rows[br].Cells[bc].Merge(c-bc, r-br)
				report.Sheets[s].Rows[br].Cells[bc]

				/*st := report.Sheets[s].Rows[br].Cells[bc].GetStyle()
				fmt.Printf("%v", st)
				st.Alignment.Horizontal = "right"
				report.Sheets[s].Rows[br].Cells[bc].SetStyle(st)*/
				//}

				b = false

			}
		}
	}
}
