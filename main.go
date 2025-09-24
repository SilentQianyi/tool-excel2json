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

		if strings.HasSuffix(filename, ".xlsx") == false {
			log.Println(fmt.Sprintf("strings.HasSuffix xlsx error! dirname[ %s ], filename[ %s ]", dirname, filename))
			continue
		}

		if strings.HasPrefix(filename, "~") {
			log.Println(fmt.Sprintf("strings.HasPrefix ~ error! dirname[ %s ], filename[ %s ]", dirname, filename))
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

	// TODO: 预处理结构问题
	exportMap := make(map[string]string)
	preDataMap := make(map[string]map[string]string)

	for _, sheet := range sheetList {

		if strings.HasPrefix(sheet, "sheet") || strings.HasPrefix(sheet, "Sheet") {
			log.Println(fmt.Sprintf("OpenFile strings.HasPrefix 2 error! dirname[ %s ], filename[ %s ], sheet[ %s ]", dirname, filename, sheet))
			continue
		}

		rows, err := f.GetRows(sheet)
		if err != nil {
			log.Fatal(fmt.Sprintf("OpenFile f.GetCols error! filename[ %s ], sheet[ %s ], err[ %s ]", filename, sheet, err.Error()))
			return err
		}

		if rows[0] == nil {
			log.Println(fmt.Sprintf("OpenFile export error! dirname[ %s ], filename[ %s ], sheet[ %s ]", dirname, filename, sheet))
			continue
		}

		if len(rows[0][0]) == 0 {
			log.Println(fmt.Sprintf("OpenFile export error 2! dirname[ %s ], filename[ %s ], sheet[ %s ]", dirname, filename, sheet))
			continue
		}

		curDataList, curMetaList, err := OpenSheet(rows)
		if err != nil {
			log.Fatal(fmt.Sprintf("OpenFile OpenSheet error! filename[ %s ], sheet[ %s ], err[ %s ]", filename, sheet, err.Error()))
			return err
		}

		outMap := toJsonStruct(curDataList, curMetaList)
		preDataMap[sheet] = outMap
		// for key, str := range outMap {
		// 	log.Println(fmt.Sprintf("OpenFile sheet[ %s ], key[ %s ], str[ %s ]", sheet, key, str))
		// }
		// log.Println(fmt.Sprintf("OpenFile sheet[ %s ] end", sheet))
	}

	for _, sheet := range sheetList {

		if strings.HasPrefix(sheet, "sheet") || strings.HasPrefix(sheet, "Sheet") {
			log.Println(fmt.Sprintf("OpenFile 2 strings.HasPrefix 2 error! dirname[ %s ], filename[ %s ], sheet[ %s ]", dirname, filename, sheet))
			continue
		}

		rows, err := f.GetRows(sheet)
		if err != nil {
			log.Fatal(fmt.Sprintf("OpenFile 2 f.GetCols error! filename[ %s ], sheet[ %s ], err[ %s ]", filename, sheet, err.Error()))
			return err
		}

		if rows[0] == nil {
			log.Println(fmt.Sprintf("OpenFile 2 export error! dirname[ %s ], filename[ %s ], sheet[ %s ]", dirname, filename, sheet))
			continue
		}

		if len(rows[0][0]) == 0 {
			log.Println(fmt.Sprintf("OpenFile 2 export error 2! dirname[ %s ], filename[ %s ], sheet[ %s ]", dirname, filename, sheet))
			continue
		}

		curDataList, curMetaList, err := OpenSheet(rows)
		if err != nil {
			log.Fatal(fmt.Sprintf("OpenFile 2 OpenSheet error! filename[ %s ], sheet[ %s ], err[ %s ]", filename, sheet, err.Error()))
			return err
		}

		if rows[0][0] == "export=server" || rows[0][0] == "export=host" {
			log.Println(fmt.Sprintf("OpenFile 2 sheet[ %s ]", sheet))
			exportMap[sheet] = toJson(curDataList, curMetaList, preDataMap)
		}
	}

	for sheet, str := range exportMap {
		output(outPath, fmt.Sprintf("%s.json", sheet), str)
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

func toJsonStruct(dataRows []rowdata, metaList []*Meta) map[string]string {
	enumValueMap := make(map[string]int)
	enumMap := make(map[string]map[string]int)
	dataMap := make(map[string]string)

	for _, row := range dataRows {
		id := ""
		line := "{"
		for idx, meta := range metaList {
			line += fmt.Sprintf("\n\t\t\t\t\"%s\": ", meta.Key)
			switch meta.Typ {
			case "string":
				if row[idx] == nil || row[idx] == "" {
					line += "\"\""
				} else {
					str := fmt.Sprintf("\"%s\"", strings.ReplaceAll(row[idx].(string), "\"", "\\\""))
					line += str
				}
			case "int":
				fallthrough
			case "uint":
				if meta.Key == "Id" {
					if row[idx] == nil || row[idx] == "" {
						break
					}
					id = fmt.Sprintf("%s", row[idx])
				}
				fallthrough
			case "float":
				if row[idx] == nil || row[idx] == "" {
					line += "0"
				} else {
					line += fmt.Sprintf("%s", row[idx])
				}
			case "Enum":
				if row[idx] == nil || row[idx] == "" {
					line += "0"
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

					line += fmt.Sprintf("%d", enumMap[key][row[idx].(string)])
				}
			case "bool":
				if row[idx] == nil || row[idx] == "" {
					line += "false"
				} else if strings.ToLower(row[idx].(string)) == "true" {
					line += "true"
				} else {
					line += "false"
				}
			case "ints":
				fallthrough
			case "uints":
				fallthrough
			case "strings":
				if row[idx] == nil || row[idx] == "" {
					line += "[]"
				} else {
					line += fmt.Sprintf("%s", row[idx])
				}
			case "Vector2":
				if row[idx] == nil || row[idx] == "" {
					line += "{\n\t\t\t\t\t\"x\": 0,\n\t\t\t\t\t\"y\": 0\n\t\t\t\t}"
				} else {
					newRow := strings.ReplaceAll(row[idx].(string), "(", "")
					newRow = strings.ReplaceAll(newRow, ")", "")
					newRowList := strings.Split(newRow, ",")
					if len(newRowList) != 2 {
						line += "{\n\t\t\t\t\t\"x\": 0,\n\t\t\t\t\t\"y\": 0\n\t\t\t\t}"
					} else {
						line += fmt.Sprintf("{\n\t\t\t\t\t\"x\": %s,\n\t\t\t\t\t\"y\": %s\n\t\t\t\t}", newRowList[0], newRowList[1])
					}
				}
			case "Vector3":
				if row[idx] == nil || row[idx] == "" {
					line += "{\n\t\t\t\t\t\"x\": 0,\n\t\t\t\t\t\"y\": 0,\n\t\t\t\t\t\"z\": 0\n\t\t\t\t}"
				} else {
					newRow := strings.ReplaceAll(row[idx].(string), "(", "")
					newRow = strings.ReplaceAll(newRow, ")", "")
					newRowList := strings.Split(newRow, ",")
					if len(newRowList) != 3 {
						line += "{\n\t\t\t\t\t\"x\": 0,\n\t\t\t\t\t\"y\": 0,\n\t\t\t\t\t\"z\": 0\n\t\t\t\t}"
					} else {
						line += fmt.Sprintf("{\n\t\t\t\t\t\"x\": %s,\n\t\t\t\t\t\"y\": %s,\n\t\t\t\t\t\"z\": %s\n\t\t\t\t}", newRowList[0], newRowList[1], newRowList[2])
					}
				}
			default:
				if row[idx] == nil || row[idx] == "" {
					line += "0"
				} else {
					line += fmt.Sprintf("%s", row[idx])
				}
			}
			line += ","
		}

		line = line[:len(line)-1]

		line += "\n\t\t\t}"
		if len(id) == 0 || id == "0" {
			break
		}
		dataMap[id] = line
	}

	return dataMap
}

func toJson(dataRows []rowdata, metaList []*Meta, preDataMap map[string]map[string]string) string {
	enumValueMap := make(map[string]int)
	enumMap := make(map[string]map[string]int)

	ret := "["
	for _, row := range dataRows {
		id := ""
		line := "\n\t{"
		for idx, meta := range metaList {
			line += fmt.Sprintf("\n\t\t\"%s\": ", meta.Key)
			switch meta.Typ {
			case "string":
				if row[idx] == nil || row[idx] == "" {
					line += "\"\""
				} else {
					line += fmt.Sprintf("\"%s\"", strings.ReplaceAll(row[idx].(string), "\"", "\\\""))
				}
			case "int":
				fallthrough
			case "uint":
				if meta.Key == "Id" {
					if row[idx] == nil || row[idx] == "" {
						break
					}
					id = fmt.Sprintf("%s", row[idx])
				}
				fallthrough
			case "float":
				if row[idx] == nil || row[idx] == "" {
					line += "0"
				} else {
					line += fmt.Sprintf("%s", row[idx])
				}
			case "Enum":
				if row[idx] == nil || row[idx] == "" {
					line += "0"
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

					line += fmt.Sprintf("%d", enumMap[key][row[idx].(string)])
				}
			case "bool":
				if row[idx] == nil || row[idx] == "" {
					line += "false"
				} else if strings.ToLower(row[idx].(string)) == "true" {
					line += "true"
				} else {
					line += "false"
				}
			case "ints":
				fallthrough
			case "uints":
				fallthrough
			case "strings":
				if row[idx] == nil || row[idx] == "" {
					line += "[]"
				} else {
					line += fmt.Sprintf("%s", row[idx])
				}
			case "Vector2":
				if row[idx] == nil || row[idx] == "" {
					line += "{\n\t\t\t\"x\": 0,\n\t\t\t\"y\": 0\n\t\t}"
				} else {
					newRow := strings.ReplaceAll(row[idx].(string), "(", "")
					newRow = strings.ReplaceAll(newRow, ")", "")
					newRowList := strings.Split(newRow, ",")
					if len(newRowList) != 2 {
						line += "{\n\t\t\t\"x\": 0,\n\t\t\t\"y\": 0\n\t\t}"
					} else {
						line += fmt.Sprintf("{\n\t\t\t\"x\": %s,\n\t\t\t\"y\": %s\n\t\t}", newRowList[0], newRowList[1])
					}
				}
			case "Vector3":
				if row[idx] == nil || row[idx] == "" {
					line += "{\n\t\t\t\"x\": 0,\n\t\t\t\"y\": 0,\n\t\t\t\"z\": 0\n\t\t}"
				} else {
					newRow := strings.ReplaceAll(row[idx].(string), "(", "")
					newRow = strings.ReplaceAll(newRow, ")", "")
					newRowList := strings.Split(newRow, ",")
					if len(newRowList) != 3 {
						line += "{\n\t\t\t\"x\": 0,\n\t\t\t\"y\": 0,\n\t\t\t\"z\": 0\n\t\t}"
					} else {
						line += fmt.Sprintf("{\n\t\t\t\"x\": %s,\n\t\t\t\"y\": %s,\n\t\t\t\"z\": %s\n\t\t}", newRowList[0], newRowList[1], newRowList[2])
					}
				}
			case "Struct":
				if row[idx] == nil || row[idx] == "" || preDataMap[meta.Key] == nil || preDataMap[meta.Key][row[idx].(string)] == "" {
					log.Fatalln(fmt.Sprintf("toJson preDataMap key[ %s ], val[ %s ]", meta.Key, row[idx].(string)))
					line += "{}"
				} else {
					log.Println(fmt.Sprintf("toJson key[ %s ], val[ %s ]", meta.Key, row[idx].(string)))
					line += preDataMap[meta.Key][row[idx].(string)]
				}
			case "StructList":
				if row[idx] == nil || row[idx] == "" {
					line += "[]"
				} else {
					str := row[idx].(string)
					str = strings.ReplaceAll(str, "[", "")
					str = strings.ReplaceAll(str, "]", "")
					strList := strings.Split(str, ",")
					line += "[\n\t\t\t"
					for _, key := range strList {
						// log.Println(fmt.Sprintf("toJson 2 key[ %s ], val[ %s ]", meta.Key, key))
						_, ok1 := preDataMap[meta.Key]
						if !ok1 {
							log.Fatalln(fmt.Sprintf("toJson preDataMap 2 key[ %s ], val[ %s ]", meta.Key, row[idx].(string)))
							line += "[]"
							break
						} else {
							_, ok2 := preDataMap[meta.Key][key]
							if !ok2 {
								log.Fatalln(fmt.Sprintf("toJson preDataMap 3 key[ %s ], val[ %s ]", meta.Key, row[idx].(string)))
								line += "[]"
								break
							} else {
								line += preDataMap[meta.Key][key]
							}
						}
						line += ",\n\t\t\t"
					}
					line = line[:len(line)-5]
					line += "\n\t\t]"
				}
			default:
				if row[idx] == nil || row[idx] == "" {
					line += "0"
				} else {
					line += fmt.Sprintf("%s", row[idx])
				}
			}
			line += ","
		}
		line = line[:len(line)-1]

		line += "\n\t},"

		if len(id) > 0 && id != "0" {
			ret += line
		}
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
