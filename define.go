package excelizer

import "github.com/xuri/excelize/v2"

type Excelizer struct {
	templates  map[string]*TemplateConfig
	tagNameMap map[string]string
}

type TemplateConfig struct {
	Name          string
	StaticSheets  []SheetConfig
	DynamicSheets []DynamicSheetConfig
	FileNameField string
}

type SheetConfig struct {
	Name   string
	Blocks []BlockConfig
}

type DynamicSheetConfig struct {
	SheetNameField   string
	Blocks           []BlockConfig
	DynamicFieldName string
}

type BlockConfig struct {
	BlockHeader    CellConfig
	Type           string
	AddChartSeries func(v interface{}) []excelize.ChartSeries
	ColumnWidths   []float64

	StaticFieldName string
	StaticCells     [][]CellConfig

	DynamicFieldName string
	DynamicCells     []CellConfig

	SubBlockFieldName string
	SubBlockConfig    *BlockConfig
}

type CellConfig struct {
	Header    string
	Value     string
	TransForm func(interface{}) interface{}
	Relatives []string
	Format    string
	Merge     MergeDirection
	Style     *CellStyle
}

type CellStyle struct {
	FontSize  float64
	FontColor string
	Border    bool
	Alignment string
}

type MergeDirection int

const (
	MergeLeft MergeDirection = iota + 1
	MergeUp
)

type Param struct {
	Req  interface{}
	Resp interface{}
}
