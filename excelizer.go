package excelizer

import (
	"fmt"
	"os"
	"path"
	"reflect"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

var excelizer *Excelizer

func NewExcelizer() *Excelizer {
	return &Excelizer{
		templates: map[string]*TemplateConfig{},
	}
}

var once sync.Once

func GetExcelizer() *Excelizer {
	once.Do(func() {
		excelizer = NewExcelizer()
	})
	return excelizer
}

func (g *Excelizer) RegisterTemplate(name string, config *TemplateConfig) {
	if config == nil {
		return
	}
	g.templates[name] = config
}

func (g *Excelizer) Generate(templateName string, data interface{}) (string, error) {
	template, exists := g.templates[templateName]
	if !exists {
		return "", fmt.Errorf("template %s not found", templateName)
	}
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()
	for _, sheet := range template.StaticSheets {
		if _, err := f.NewSheet(sheet.Name); err != nil {
			return "", err
		}
		for _, block := range sheet.Blocks {
			rows, err := f.GetRows(sheet.Name)
			if err != nil {
				return "", err
			}
			rowsLen := len(rows)
			row := 1
			if rowsLen != 0 {
				row = rowsLen + 2
			}
			if err = g.generateBlock(f, sheet.Name, block, data, &row); err != nil {
				return "", err
			}
			if err = g.setColumnWidths(f, sheet.Name, &block); err != nil {
				return "", err
			}
		}
	}
	for _, dynamicSheet := range template.DynamicSheets {
		dynamicField, err := g.getFieldValue(data, dynamicSheet.DynamicFieldName)
		if err != nil {
			return "", fmt.Errorf("failed to get dynamic field: %s, err: %v", dynamicSheet.DynamicFieldName, err)
		}
		dynamicValue := reflect.ValueOf(dynamicField)
		if dynamicValue.Kind() != reflect.Slice {
			dynamicValue = reflect.ValueOf([]interface{}{dynamicValue})
		}
		for i := 0; i < dynamicValue.Len(); i++ {
			item := dynamicValue.Index(i).Interface()
			sheetName, err := g.getFieldValue(item, dynamicSheet.SheetNameField)
			if err != nil {
				return "", fmt.Errorf("failed to get sheet name field: %s, err: %v", dynamicSheet.SheetNameField, err)
			}
			sheetNameStr := fmt.Sprintf("%v", sheetName)
			if _, err := f.NewSheet(sheetNameStr); err != nil {
				return "", err
			}

			for _, block := range dynamicSheet.Blocks {
				rows, err := f.GetRows(sheetNameStr)
				if err != nil {
					return "", err
				}
				rowsLen := len(rows)
				row := 1
				if rowsLen != 0 {
					row = rowsLen + 2
				}
				if err = g.generateBlock(f, sheetNameStr, block, item, &row); err != nil {
					return "", err
				}
				if err = g.setColumnWidths(f, sheetNameStr, &block); err != nil {
					return "", err
				}
			}
		}
	}
	if sheetIndex, err := f.GetSheetIndex("Sheet1"); err == nil && sheetIndex != -1 {
		if rows, _ := f.GetRows("Sheet1"); len(rows) == 0 {
			_ = f.DeleteSheet("Sheet1")
		}
	}
	fileName := templateName
	if template.FileNameField != "" {
		if parsedFileName, err := g.getFieldValue(data, template.FileNameField); err == nil && parsedFileName != nil && fmt.Sprintf("%v", parsedFileName) != "" {
			fileName = fmt.Sprintf("%v", parsedFileName)
		}
	}
	var absolutePath string
	loc, _ := time.LoadLocation("Asia/Shanghai")
	absolutePath = path.Join("/tmp", time.Now().In(loc).Format("20060102150405"), fileName+".xlsx")
	if err := os.MkdirAll(path.Dir(absolutePath), os.ModePerm); err != nil {
		return "", err
	}
	return absolutePath, f.SaveAs(absolutePath)
}

const (
	StaticCell   string = "staticCell"
	DynamicCell  string = "dynamicCell"
	DynamicBlock string = "dynamicBlock"
)

func (g *Excelizer) generateBlock(f *excelize.File, sheetName string, blockConfig BlockConfig, data interface{}, row *int) (err error) {
	if blockConfig.BlockHeader.Header != "" || blockConfig.BlockHeader.Value != "" {
		if err = g.applyCellConfig(f, sheetName, blockConfig.BlockHeader, fmt.Sprintf("A%d", *row), data); err != nil {
			fmt.Printf("block header err: %v\n", err)
			return err
		}
		*row++
	}
	field := data
	switch blockConfig.Type {
	case StaticCell:
		if field, err = g.getFieldForBlockType(data, blockConfig.StaticFieldName); err != nil {
			return fmt.Errorf("unknown field: %s, err: %+v", blockConfig.StaticFieldName, err.Error())
		}
		return g.generateStaticCell(f, sheetName, blockConfig, field, row)
	case DynamicCell:
		if field, err = g.getFieldForBlockType(data, blockConfig.DynamicFieldName); err != nil {
			return fmt.Errorf("unknown field: %s, err: %+v", blockConfig.DynamicFieldName, err.Error())
		}
		if err = g.generateDynamicCell(f, sheetName, blockConfig, field, row); err != nil {
			return err
		}
		if len(blockConfig.StaticCells) > 0 {
			if field, err = g.getFieldForBlockType(data, blockConfig.StaticFieldName); err != nil {
				return fmt.Errorf("unknown field: %s, err: %+v", blockConfig.StaticFieldName, err.Error())
			}
			return g.generateStaticCell(f, sheetName, blockConfig, field, row)
		}
	case DynamicBlock:
		if blockConfig.SubBlockConfig == nil {
			return fmt.Errorf("unknown subBlockConfig")
		}
		if field, err = g.getFieldForBlockType(data, blockConfig.SubBlockFieldName); err != nil {
			return fmt.Errorf("unknown field: %s, err: %+v", blockConfig.SubBlockFieldName, err.Error())
		}
		if err = g.generateDynamicBlock(f, sheetName, *blockConfig.SubBlockConfig, field, row); err != nil {
			return err
		}
		if len(blockConfig.StaticCells) > 0 {
			if field, err = g.getFieldForBlockType(data, blockConfig.StaticFieldName); err != nil {
				return fmt.Errorf("unknown field: %s, err: %+v", blockConfig.StaticFieldName, err.Error())
			}
			return g.generateStaticCell(f, sheetName, blockConfig, field, row)
		}
	default:
		return fmt.Errorf("unknown block type: %s", blockConfig.Type)
	}
	return nil
}

func (g *Excelizer) getFieldForBlockType(data interface{}, fieldName string) (interface{}, error) {
	if fieldName == "" {
		return data, nil
	}
	return g.getFieldValue(data, fieldName)
}

func (g *Excelizer) generateStaticCell(f *excelize.File, sheetName string, blockConfig BlockConfig, data interface{}, row *int) error {
	for _, rowConfig := range blockConfig.StaticCells {
		for colIndex, cellConfig := range rowConfig {
			cellName := fmt.Sprintf("%c%d", 'A'+colIndex, *row)
			if err := g.applyCellConfig(f, sheetName, cellConfig, cellName, data); err != nil {
				return err
			}
		}
		*row++
	}
	return nil
}

func (g *Excelizer) generateDynamicCell(f *excelize.File, sheetName string, blockConfig BlockConfig, data interface{}, row *int) error {
	for i, cellConfig := range blockConfig.DynamicCells {
		cellName := fmt.Sprintf("%c%d", 'A'+i, *row)
		_ = f.SetCellValue(sheetName, cellName, cellConfig.Header)
		style := cellConfig.Style
		if style == nil {
			style = &CellStyle{FontSize: 12, FontColor: "#000000", Border: true}
		}
		excelStyle, err := g.createExcelStyle(f, style)
		if err != nil {
			return err
		}
		_ = f.SetCellStyle(sheetName, cellName, cellName, excelStyle)
	}
	*row++

	value := reflect.ValueOf(data)
	if value.Kind() != reflect.Slice {
		value = reflect.ValueOf([]interface{}{data})
	}
	if value.Kind() == reflect.Slice {
		slice := value
		for rowIndex := 0; rowIndex < slice.Len(); rowIndex++ {
			item := slice.Index(rowIndex).Interface()
			for colIndex, cellConfig := range blockConfig.DynamicCells {
				cellName := fmt.Sprintf("%c%d", 'A'+colIndex, *row)
				if err := g.applyCellConfig(f, sheetName, cellConfig, cellName, item); err != nil {
					return err
				}
			}
			*row++
		}
	}

	if blockConfig.AddChartSeries != nil {
		_ = f.AddChart(sheetName, fmt.Sprintf("%c%d", 'A'+len(blockConfig.DynamicCells)+1, *row), &excelize.Chart{
			Type:   excelize.Col3D,
			Series: blockConfig.AddChartSeries(data),
			Legend: excelize.ChartLegend{
				Position: "right",
			},
			PlotArea: excelize.ChartPlotArea{
				ShowCatName:     true,
				ShowLeaderLines: false,
				ShowPercent:     false,
				ShowSerName:     true,
				ShowVal:         true,
			},
		})
		*row++
	}
	return nil
}

func (g *Excelizer) generateDynamicBlock(f *excelize.File, sheetName string, blockConfig BlockConfig, data interface{}, row *int) error {
	value := reflect.ValueOf(data)
	if value.Kind() != reflect.Slice {
		value = reflect.ValueOf([]interface{}{data})
	}
	if value.Kind() == reflect.Slice {
		slice := value
		for rowIndex := 0; rowIndex < slice.Len(); rowIndex++ {
			item := slice.Index(rowIndex).Interface()
			if err := g.generateBlock(f, sheetName, blockConfig, item, row); err != nil {
				return err
			}
			*row += 2 // 每个block之间留白两行
		}
	}
	return nil
}

func (g *Excelizer) setColumnWidths(f *excelize.File, sheetName string, block *BlockConfig) error {
	if block.ColumnWidths != nil {
		for i, width := range block.ColumnWidths {
			col := fmt.Sprintf("%c", 'A'+i)
			if err := f.SetColWidth(sheetName, col, col, width); err != nil {
				return err
			}
		}
	} else {
		defaultWidth := float64(20)
		for i := 0; i < GetColumnCount(f, sheetName); i++ {
			col := fmt.Sprintf("%c", 'A'+i)
			if err := f.SetColWidth(sheetName, col, col, defaultWidth); err != nil {
				return err
			}
		}
	}
	return nil
}
