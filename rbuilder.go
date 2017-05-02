package rbuilder

import (
	"bytes"
	"errors"
	"fmt"
	"strconv"
	"strings"
	"time"

	"text/template"

	"github.com/tealeg/xlsx"
)

const debug bool = true

type Template struct {
	tmpl       *xlsx.File
	staticData interface{}
}

func NewTemplate(tmpl *xlsx.File, staticData interface{}) Template {
	return Template{tmpl: tmpl, staticData: staticData}
}

// Render generates report based on template. Returns new object xlsx what
// inherits template with values instead of text/template placeholders.
func (t *Template) Render(data interface{}) (*xlsx.File, error) {

	// create template copy
	buf := bytes.NewBuffer(nil)
	err := t.tmpl.Write(buf)
	if err != nil {
		return nil, err
	}

	result, err := xlsx.OpenBinary(buf.Bytes())
	if err != nil {
		return nil, err
	}

	// render static template values {{.Attr}}, what does not
	// change amount of lines in result file
	t.renderStatic(result, data)

	// render {{range }}{{end}} what changes amount of line.
	err = t.renderRange(result, data)

	return result, err
}

func (t *Template) renderRange(report *xlsx.File, data interface{}) error {
	tags := ""

	for s := range report.Sheets {
		for r := range report.Sheets[s].Rows {
		rows:
			for c := range report.Sheets[s].Rows[r].Cells {
				val := report.Sheets[s].Rows[r].Cells[c].Value
				if strings.Contains(val, "{{") && strings.Contains(val, "}}") {

					if strings.Contains(val, "range") {
						// добавляем заголовок блока range
						tags = tags + fmt.Sprintf("##begin:%d%s", r, tagSeparator)
					}

					if strings.Contains(val, "{{end.}}") {
						val = strings.Replace(val, "{{end.}}", "", 1)
						if len(val) > 0 {
							// если в ячейке есть какие то другие данные кроме тэга {{end}}
							tags += fmt.Sprintf("%s<<%d>>", val, c)
						}

						tags += tagSeparator + "{{end}}##end" + tagSeparator

						break rows
					}
					tags = tags + fmt.Sprintf("%s<<%d>>", val, c)

				}

			}
		}
	}

	funcMap := template.FuncMap{

		"fdate": func(s string, t time.Time) string { return t.Format(s) },
	}

	debugf("%s\n", tags)
	tmp := template.Must(template.New("report").Funcs(funcMap).Parse(tags))

	buf := bytes.NewBuffer(nil)

	err := tmp.Execute(buf, struct {
		D interface{}
		S interface{}
	}{data, t.staticData})
	if err != nil {
		return err
	}

	lines := strings.Split(buf.String(), tagSeparator)

	if debug {
		for i := range lines {
			debugf("%s\n", lines[i])
		}
	}

	i := 0
	offset := 0
	for ; i < len(lines); i++ {

		debugf("%s\n", lines[i])
		if strings.HasPrefix(lines[i], "##begin:") {
			// rangeRowNum хранит номер строки из шаблона где находится {{range}}{{end}}
			rangeRowNum, err := strconv.Atoi(lines[i][8:])
			if err != nil {
				return err
			}

			cnt := 0
			for j := i + 1; j < len(lines); j++ {
				if strings.HasPrefix(lines[j], "##end") {
					cnt = j - i - 1
					break
				}
			}

			if cnt == 0 {
				debugf("del row: %d, offset:%d\n", rangeRowNum, offset)
				// значит надо удалить строчку с {{range}}
				err = delRow(report, 0, rangeRowNum+offset)
				if err != nil {
					return err
				}
				offset--
				continue
			}

			debugf("amount of rows to insert %d\n", cnt)

			// копируем все строки до начала строки range, потому что если не будет данных
			// нам не надо создавать пустую строчку без данных
			if err := insertRows(report, 0, rangeRowNum+offset, cnt-1, report.Sheets[0].Rows[rangeRowNum+offset]); err != nil {
				return err
			}

			offset += cnt - 1

			/*		// двигаемся до конца сгенерированного блока
					i++
					for ; i < len(lines); i++ {
						if strings.HasPrefix(lines[i], "##end") {

							break
						}
						// если есть строка которую надо добавить как результат range
						// добавляем одну строку
						debugf("append table row: %d\n", i)
						/*				if err := appendRows(report, result, 0, rangeRowNum, rangeRowNum+1, 0); err != nil {

										return nil, err
									}*/

			// надо ее распарсить и присворить значениям новой добавленной строки
			//}

		}
	}

	return err
}

const tagSeparator = "$$^~^$$"

func (t *Template) renderStatic(report *xlsx.File, data interface{}) error {

	// tags holds all static values what does not part of
	// {{range}}{{end}} block
	// at the moment supported only single row {{range}}..{{end}}
	// what covers whole row
	tags := ""

	if len(report.Sheets) == 0 {
		return errors.New("report has not scheets")
	}
	for s := range report.Sheets {
		for r := range report.Sheets[s].Rows {
		rows:
			for c := range report.Sheets[s].Rows[r].Cells {
				val := report.Sheets[s].Rows[r].Cells[c].Value
				if !(strings.Contains(val, "{{") && strings.Contains(val, "}}")) {
					continue
				}

				if strings.Contains(val, "range") {
					debugf("range found: %d:%d:%d\n", s, r, c)
					break rows
				}

				tags = tags + fmt.Sprintf("##%d:%d:%d##%s%s", s, r, c, val, tagSeparator)

			}
		}
	}

	debugf("Static tags positions:\n %s", tags)

	tmp := template.Must(template.New("report").Parse(tags))

	buf := bytes.NewBuffer(nil)

	err := tmp.Execute(buf, struct {
		D interface{}
		S interface{}
	}{data, t.staticData})
	if err != nil {
		return err
	}

	lines := strings.Split(buf.String(), tagSeparator)
	for _, line := range lines {
		debugf("extracting parsing results: %s\n", line)
		if !strings.HasPrefix(line, "##") {
			continue
		}

		line = line[2:]
		closeIdx := strings.Index(line, "##")
		if closeIdx < 0 {
			return errors.New("fatal error: not found closing ##")
		}

		s, r, c, err := extractSRC(line[:closeIdx])
		if err != nil {
			return err
		}

		str := line[closeIdx+2:] //##0:1:1##Привет {{.Name}}

		numberFormat := report.Sheets[0].Cell(r, c).GetNumberFormat()
		if val, err := strconv.ParseFloat(str, 10); err == nil {
			report.Sheets[s].Cell(r, c).SetFloat(val)
			report.Sheets[s].Cell(r, c).NumFmt = numberFormat
			continue
		}

		if val, err := strconv.ParseInt(str, 10, 64); err == nil {
			report.Sheets[s].Cell(r, c).SetInt64(val)
			report.Sheets[s].Cell(r, c).NumFmt = numberFormat
			continue
		}

		if report.Sheets[s].Cell(r, c).Type() == xlsx.CellTypeDate {
			println("date type cell")
			//report.Sheets[0].Cell(r, c).SetDate(time.Now())
			continue
		}

		report.Sheets[s].Cell(r, c).SetString(str)

		/*
			//style := report.Sheets[0].Cell(r, c).Set
			cellType := report.Sheets[0].Cell(r, c).Type()
			//println("format", format)
			if cellType == xlsx.CellTypeNumeric {
				fval, err := strconv.ParseFloat(val, 10)
				if err != nil {
					return err
				}
			} else {

			}
		*/
		//
	}

	return nil
}

func parseRangeLine(s string, m map[int]string) {

}

// extractSRC takes string like "1:4:5" and extracts values
func extractSRC(str string) (s, r, c int, err error) {

	val := strings.SplitN(str, ":", 3)
	if len(val) != 3 {
		err = errors.New("invalid input! Expected line look like \"1:2:5\"")
		return
	}

	s, err = strconv.Atoi(val[0])
	if err != nil {
		return
	}

	r, err = strconv.Atoi(val[1])
	if err != nil {
		return
	}

	c, err = strconv.Atoi(val[2])
	if err != nil {
		return
	}

	return
}

func debugf(s string, args ...interface{}) {
	if debug {
		fmt.Printf(s, args...)
	}
}

func appendRows(from, to *xlsx.File, fromS, fromR, toR, toS int) error {

	if from == nil {
		return errors.New("invalid source file")
	}

	if to == nil {
		return errors.New("invalid target file")
	}

	if fromS >= len(from.Sheets) {
		return errors.New("invalid scheet number from")
	}

	if toS >= len(to.Sheets) {
		return errors.New("invalid scheet number to")
	}

	if fromR >= len(from.Sheets[fromS].Rows) {
		return errors.New("invalid starting row in source scheet")
	}

	if toR >= len(from.Sheets[fromS].Rows) {
		return errors.New("invalid ending row in source scheet")
	}
	/*
		for i := fromR; i <= toR; i++ {
			row := to.Sheets[toS].AddRow()
			for _, c := range from.Sheets[fromS].Rows[i].Cells {
				cell := row.AddCell()
				cell.SetStyle(c.GetStyle())
				cell.SetValue(c.Value)
			}
		}
	*/

	for i := fromR; i <= toR; i++ {
		(*to).Sheets[toS].Rows = append((*to).Sheets[toS].Rows, from.Sheets[fromS].Rows[i])
	}
	for i := range to.Sheets[toS].Rows {
		(*to).Sheets[toS].Rows[i].Sheet = to.Sheets[toS]
	}

	return nil
}

func delRow(f *xlsx.File, s, r int) error {

	if f == nil {
		return errors.New("invalid file")
	}

	if s >= len(f.Sheets) {
		return errors.New("invalid scheet number")
	}

	if r >= len(f.Sheets[s].Rows) {
		return errors.New("invalid row in scheet")
	}

	(*f).Sheets[s].Rows = append(f.Sheets[s].Rows[0:r], append([]*xlsx.Row{}, (*f).Sheets[s].Rows[r+1:]...)...)

	return nil
}

func insertRows(f *xlsx.File, toS, startR, cnt int, row *xlsx.Row) error {

	if f == nil {
		return errors.New("invalid file")
	}

	if row == nil {
		return errors.New("invalid row")
	}

	if toS >= len(f.Sheets) {
		return errors.New("invalid scheet number")
	}

	if startR >= len(f.Sheets[toS].Rows) {
		return errors.New("invalid starting row in scheet")
	}

	for i := 0; i < cnt; i++ {
		(*f).Sheets[toS].Rows = append(f.Sheets[toS].Rows[0:startR], append([]*xlsx.Row{row}, (*f).Sheets[toS].Rows[startR:]...)...)
	}

	return nil
}
