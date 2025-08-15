package excelizer

import (
	"fmt"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func (g *Excelizer) getDefaultStyle(cellStyle *CellStyle, defaultStyle *CellStyle) *CellStyle {
	if cellStyle == nil {
		return defaultStyle
	}
	return cellStyle
}

func (g *Excelizer) applyCellConfig(f *excelize.File, sheet string, config CellConfig, cellName string, data interface{}) error {
	cellValue := config.Header
	if data != nil {
		value, err := g.getFieldValue(data, config.Value)
		if err == nil {
			if config.TransForm != nil {
				value = config.TransForm(value)
			}
			var relatedValues []interface{}
			for _, field := range config.Relatives {
				rv, err := g.getFieldValue(data, field)
				if err != nil {
					return err
				}
				relatedValues = append(relatedValues, rv)
			}
			var formattedValue string
			if config.Format != "" {
				allValues := append([]interface{}{value}, relatedValues...)
				formattedValue = g.formatWithCalculation(config.Format, allValues)
			} else {
				formattedValue = fmt.Sprintf("%v", value)
			}
			cellValue = formattedValue
		}
	}
	_ = f.SetCellValue(sheet, cellName, cellValue)
	style := config.Style
	if style == nil {
		style = &CellStyle{FontSize: 12, FontColor: "#000000", Border: true}
	}
	excelStyle, err := g.createExcelStyle(f, style)
	if err != nil {
		return err
	}
	if cellValue != "" {
		_ = f.SetCellStyle(sheet, cellName, cellName, excelStyle)
	}
	if config.Merge != 0 {
		_ = MergeCellWithDirection(f, sheet, cellName, config.Merge)
	}
	return nil
}

func (g *Excelizer) formatWithCalculation(format string, values []interface{}) string {
	// 暂只支持除法
	if strings.Contains(format, "/") {
		parts := strings.Split(format, "/")
		if len(parts) == 2 {
			num1, ok1 := values[0].(int64)
			num2, ok2 := values[1].(int64)
			if ok1 && ok2 {
				result := float64(num1) / float64(num2) * 100
				return fmt.Sprintf("%.2f%%", result)
			}
		}
	}
	return fmt.Sprintf(format, values...)
}
func MergeCellWithDirection(f *excelize.File, sheet, cell string, direction MergeDirection) error {
	x, y, err := excelize.CellNameToCoordinates(cell)
	if err != nil {
		return err
	}
	var startCell, endCell string
	switch direction {
	case MergeLeft:
		for x > 0 {
			x--
			startCell, err = excelize.CoordinatesToCellName(x, y)
			if err != nil {
				return err
			}
			value, err := f.GetCellValue(sheet, startCell)
			if err != nil {
				return err
			}
			if value != "" {
				break
			}
		}
		endCell = cell
	case MergeUp:
		for y > 0 {
			y--
			startCell, err = excelize.CoordinatesToCellName(x, y)
			if err != nil {
				return err
			}
			value, err := f.GetCellValue(sheet, startCell)
			if err != nil {
				return err
			}
			if value != "" {
				break
			}
		}
		endCell = cell
	default:
		return fmt.Errorf("unsupported direction: %d", direction)
	}
	if startCell != "" && endCell != "" {
		if err = f.MergeCell(sheet, startCell, endCell); err != nil {
			return err
		}
		_ = f.SetCellStyle(sheet, startCell, endCell, 1)
	}
	return nil
}

func (g *Excelizer) createExcelStyle(f *excelize.File, style *CellStyle) (int, error) {
	styleConfig := &excelize.Style{
		Font: &excelize.Font{
			Size:  style.FontSize,
			Color: style.FontColor,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
			WrapText:   true,
		},
	}
	if style.Border {
		styleConfig.Border = []excelize.Border{
			{Type: "left", Color: "#000000", Style: 1},
			{Type: "top", Color: "#000000", Style: 1},
			{Type: "bottom", Color: "#000000", Style: 1},
			{Type: "right", Color: "#000000", Style: 1},
		}
	}
	return f.NewStyle(styleConfig)
}

func GetColumnCount(f *excelize.File, sheetName string) int {
	rows, err := f.GetRows(sheetName)
	if err != nil || len(rows) == 0 {
		return 0
	}
	maxCol := 0
	for i := 0; i < len(rows); i++ {
		curCol := len(rows[i])
		if curCol > maxCol {
			maxCol = curCol
		}
	}
	return maxCol
}

func (g *Excelizer) getFieldValue(data interface{}, fieldPath string) (interface{}, error) {
	if fieldPath == "" {
		return data, fmt.Errorf("no specified fieldPath")
	}
	value := reflect.ValueOf(data)
	fields := strings.Split(fieldPath, ".")
	for _, field := range fields {
		for value.Kind() == reflect.Ptr || value.Kind() == reflect.Interface {
			value = value.Elem()
		}
		switch value.Kind() {
		case reflect.Struct:
			value = value.FieldByName(field)
			if !value.IsValid() {
				return nil, fmt.Errorf("field not found: %s", field)
			}
		case reflect.Map:
			key := reflect.ValueOf(field)
			value = value.MapIndex(key)
			if value.IsZero() {
				return nil, fmt.Errorf("key not found: %s", field)
			}
		case reflect.Slice, reflect.Array:
			index, err := strconv.Atoi(field)
			if err != nil {
				return nil, fmt.Errorf("invalid index: %s", field)
			}
			if index < 0 || index >= value.Len() {
				return nil, fmt.Errorf("index out of range: %d", index)
			}
			value = value.Index(index)
		default:
			return nil, fmt.Errorf("invalid kind: %s", value.Kind())
		}
	}
	return value.Interface(), nil
}

func TransformBytes(v interface{}) interface{} {
	bytes, ok := v.(int64)
	if !ok {
		return v
	}
	if bytes == 0 {
		return "0 B"
	}
	units := []string{"B", "KB", "MB", "GB", "TB", "PB"}
	var unitIndex int
	fBytes := float64(bytes)
	for fBytes >= 1024 && unitIndex < len(units)-1 {
		fBytes /= 1024
		unitIndex++
	}
	return fmt.Sprintf("%.2f %s", fBytes, units[unitIndex])
}

func TransFormTime(v interface{}) interface{} {
	t, ok := v.(time.Time)
	if !ok {
		return v
	}
	return t.Format("2006-01-02 15:04:05")
}
