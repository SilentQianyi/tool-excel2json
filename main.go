package main

import (
	"errors"
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	args := os.Args
	if len(args) < 3 {
		log.Fatal("insufficient args ", len(args))
	}

	dirname := args[1]
	outPath := args[2]

	entries, err := os.ReadDir(dirname)
	if err != nil {
		log.Fatal(fmt.Sprintf("os.ReadDir error! dirname[ %s ], err[ %s ]", dirname, err.Error()))
	}
	filenames := make([]string, 0, len(entries))
	for _, entry := range entries {
		info, err := entry.Info()
		if err != nil {
			log.Fatal(fmt.Sprintf("entry.Info() error! dirname[ %s ], err[ %s ]", dirname, err.Error()))
			continue
		}
		filenames = append(filenames, info.Name())
	}

	for _, filename := range filenames {

		if strings.HasPrefix(filename, "~") {
			log.Println(fmt.Sprintf("strings.HasPrefix error! dirname[ %s ], filename[ %s ]", dirname, filename))
			continue
		}

		err = OpenFile(dirname, filename, outPath)
		if err != nil {
			log.Fatal(fmt.Sprintf("OpenFile error! dirname[ %s ], filename[ %s ], err[ %s ]", dirname, filename, err.Error()))
		}
	}
}

func OpenFile(dirname string, filename string, outPath string) (err error) {
	f, err := excelize.OpenFile(dirname + "/" + filename)
	if err != nil {
		log.Fatal(fmt.Sprintf("OpenFile excelize.OpenFile error! dirname[ %s ], filename[ %s ], err[ %s ]", dirname, filename, err.Error()))
		return err
	}

	defer func() {
		// Close the spreadsheet.
		if err = f.Close(); err != nil {
			log.Fatal(fmt.Sprintf("OpenFile f.Close() error! dirname[ %s ], filename[ %s ], err[ %s ]", dirname, filename, err.Error()))
		}
	}()

	sheetList := f.GetSheetList()

	for _, sheet := range sheetList {

		if strings.HasPrefix(sheet, "sheet") || strings.HasPrefix(sheet, "Sheet") {
			log.Println(fmt.Sprintf("strings.HasPrefix 2 error! dirname[ %s ], filename[ %s ], sheet[ %s ]", dirname, filename, sheet))
			continue
		}

		rows, err := f.GetRows(sheet)
		if err != nil {
			log.Fatal(fmt.Sprintf("OpenFile f.GetCols error! filename[ %s ], sheet[ %s ], err[ %s ]", filename, sheet, err.Error()))
			return err
		}

		if rows[0] == nil || (rows[0][0] != "export=server" && rows[0][0] != "export=host") {
			log.Println(fmt.Sprintf("export error! dirname[ %s ], filename[ %s ], sheet[ %s ]", dirname, filename, sheet))
			continue
		}

		curDataList, curMetaList, err := OpenSheet(rows)
		if err != nil {
			log.Fatal(fmt.Sprintf("OpenFile OpenSheet error! filename[ %s ], sheet[ %s ], err[ %s ]", filename, sheet, err.Error()))
			return err
		}

		output(outPath, fmt.Sprintf("%s.json", sheet), toJson(curDataList, curMetaList))
	}

	return nil
}

type Meta struct {
	Key string
	Idx int
	Typ string
}

type rowdata []interface{}

func OpenSheet(rows [][]string) (dataList []rowdata, metaList []*Meta, err error) {

	rowLen := len(rows)
	if rowLen < 4 {
		log.Fatal(fmt.Sprintf("OpenSheet row len error! err[ %s ]", err.Error()))
		return nil, nil, errors.New("OpenSheet row len error")
	}

	colNum := len(rows[1])
	metaList = make([]*Meta, colNum)
	dataList = make([]rowdata, len(rows)-4)

	for rowIndex, row := range rows {
		switch rowIndex {
		case 0:
		case 1:
			for idx, typ := range row {
				metaList[idx] = &Meta{
					Key: "",
					Idx: idx,
					Typ: typ,
				}
			}
		case 2:
			for idx, name := range row {
				metaList[idx].Key = name
			}
		case 3:
		default:
			data := make(rowdata, colNum)

			for i := 0; i < colNum; i++ {
				if i < len(row) {
					data[i] = row[i]
				}
			}

			dataList[rowIndex-4] = data
		}
	}

	return dataList, metaList, nil
}

func toJson(dataRows []rowdata, metaList []*Meta) string {
	enumValueMap := make(map[string]int)
	enumMap := make(map[string]map[string]int)

	ret := "["
	for _, row := range dataRows {
		ret += "\n\t{"
		for idx, meta := range metaList {
			ret += fmt.Sprintf("\n\t\t\"%s\": ", meta.Key)
			switch meta.Typ {
			case "string":
				if row[idx] == nil || row[idx] == "" {
					ret += "\"\""
				} else {
					ret += fmt.Sprintf("\"%s\"", strings.ReplaceAll(row[idx].(string), "\"", "\\\""))
				}
			case "int":
				fallthrough
			case "uint":
				fallthrough
			case "float":
				if row[idx] == nil || row[idx] == "" {
					ret += "0"
				} else {
					ret += fmt.Sprintf("%s", row[idx])
				}
			case "Enum":
				if row[idx] == nil || row[idx] == "" {
					ret += "0"
				} else {
					key := meta.Key
					_, ok := enumMap[key]
					if !ok {
						enumMap[key] = make(map[string]int)
						enumValueMap[key] = 0
					}
					_, ok = enumMap[key][row[idx].(string)]
					if !ok {
						enumMap[key][row[idx].(string)] = enumValueMap[key]
						enumValueMap[key]++
					}

					ret += fmt.Sprintf("%d", enumMap[key][row[idx].(string)])
				}
			case "bool":
				if row[idx] == nil || row[idx] == "" {
					ret += "false"
				} else if strings.ToLower(row[idx].(string)) == "true" {
					ret += "true"
				} else {
					ret += "false"
				}
			case "ints":
				fallthrough
			case "uints":
				fallthrough
			case "strings":
				if row[idx] == nil || row[idx] == "" {
					ret += "[]"
				} else {
					ret += fmt.Sprintf("[%s]", row[idx])
				}
			case "Vector2":
				if row[idx] == nil || row[idx] == "" {
					ret += "{\n\t\t\t\"x\": 0,\n\t\t\t\"y\": 0\n\t\t}"
				} else {
					newRow := strings.ReplaceAll(row[idx].(string), "(", "")
					newRow = strings.ReplaceAll(newRow, ")", "")
					newRowList := strings.Split(newRow, ",")
					if len(newRowList) != 2 {
						ret += "{\n\t\t\t\"x\": 0,\n\t\t\t\"y\": 0\n\t\t}"
					} else {
						ret += fmt.Sprintf("{\n\t\t\t\"x\": %s,\n\t\t\t\"y\": %s\n\t\t}", newRowList[0], newRowList[1])
					}
				}
			case "Vector3":
				if row[idx] == nil || row[idx] == "" {
					ret += "{\n\t\t\t\"x\": 0,\n\t\t\t\"y\": 0,\n\t\t\t\"z\": 0\n\t\t}"
				} else {
					newRow := strings.ReplaceAll(row[idx].(string), "(", "")
					newRow = strings.ReplaceAll(newRow, ")", "")
					newRowList := strings.Split(newRow, ",")
					if len(newRowList) != 3 {
						ret += "{\n\t\t\t\"x\": 0,\n\t\t\t\"y\": 0,\n\t\t\t\"z\": 0\n\t\t}"
					} else {
						ret += fmt.Sprintf("{\n\t\t\t\"x\": %s,\n\t\t\t\"y\": %s,\n\t\t\t\"z\": %s\n\t\t}", newRowList[0], newRowList[1], newRowList[2])
					}
				}
			default:
				if row[idx] == nil || row[idx] == "" {
					ret += "0"
				} else {
					ret += fmt.Sprintf("%s", row[idx])
				}
			}
			ret += ","
		}
		ret = ret[:len(ret)-1]

		ret += "\n\t},"
	}
	ret = ret[:len(ret)-1]

	ret += "\n]"
	return ret
}

func output(outPath string, filename string, str string) error {

	f, err := os.OpenFile(outPath+"/"+strings.ToLower(filename), os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0777)
	if err != nil {
		return err
	}
	defer f.Close()

	_, err = f.WriteString(str)
	if err != nil {
		return err
	}

	return nil
}
