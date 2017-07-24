package rbuilder_test

import (
	"testing"

	"time"

	"github.com/regorov/rbuilder"
	"github.com/tealeg/xlsx"
)

func TestCloneSheet(t *testing.T) {

	f, err := xlsx.OpenFile("ko-pay-note-test.xlsx")
	if err != nil {
		t.Fatal(err)
	}

	d := map[string]interface{}{"License": "Лицензия 123-456-78 от 27.09.13 г.",
		"CompanyName": "ООО \"Элефант софт\" г.Казань",
		"CurrentTime": time.Now().Format("02.01.2006 15:04:05"),
		"Dyn0":        []int{111, 222, 333, 444},
		"Dyn1":        []int{555, 666, 777, 888},
	}

	err = rbuilder.CloneSheet(f, 1, "Отрезок-1", "Dyn0", "Dyn1")
	if err != nil {
		t.Error(err)
	}

	err = f.Save("test1.xlsx")
	if err != nil {
		t.Error(err)
	}

	tmpl := rbuilder.NewTemplate(f, nil)

	out, err := tmpl.Render(d)
	if err != nil {
		t.Error(err)
	}

	err = out.Save("test1-1.xlsx")
	if err != nil {
		t.Error(err)
	}

}
