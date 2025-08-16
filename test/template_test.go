package test

import (
	"excelizer"
	"testing"
	"time"
)

var excelGenerator = excelizer.GetExcelizer()

const StaticTemplate string = "StaticTemplate"

func TestStaticExcel(t *testing.T) {
	Req := struct {
		StartTime time.Time `json:"startTime"`
		EndTime   time.Time `json:"endTime"`
	}{
		StartTime: time.Now(),
		EndTime:   time.Now(),
	}
	resp := struct {
		Header string `json:"header"`
		Value  string `json:"value"`
		Data   int32  `json:"data"`
	}{
		Header: "Header",
		Value:  "test",
		Data:   1024,
	}
	excelGenerator.RegisterTemplate(StaticTemplate, &excelizer.TemplateConfig{
		Name: StaticTemplate,
		StaticSheets: []excelizer.SheetConfig{
			{
				Name: "sheet1",
				Blocks: []excelizer.BlockConfig{
					{
						Type:            excelizer.StaticCell,
						StaticFieldName: "Req",
						StaticCells: [][]excelizer.CellConfig{
							{
								{Header: "TimeRange"},
								{Value: "StartTime", Format: "%v", TransForm: excelizer.TransFormTime},
								{Header: "-"},
								{Value: "EndTime", Format: "%v", TransForm: excelizer.TransFormTime},
							},
						},
					},
					{
						Type:             excelizer.DynamicCell,
						BlockHeader:      excelizer.CellConfig{Value: "Resp.Header"},
						DynamicFieldName: "Resp",
						DynamicCells: []excelizer.CellConfig{
							{Header: "", Value: "Value"},
							{Header: "Data", Value: "Data", TransForm: excelizer.TransformBytes},
						},
					},
				},
			},
		},
	})
	_, _ = excelGenerator.Generate(StaticTemplate, excelizer.Param{Req: Req, Resp: resp})
}

const SubBlockTemplate string = "SubBlockTemplate"

func TestSubBlockExcel(t *testing.T) {
	type SubInfo struct {
		Value1 string `json:"value1"`
		Value2 string `json:"value2"`
	}
	resp := struct {
		SheetName string    `json:"sheetName"`
		Sub       []SubInfo `json:"sub"`
	}{
		SheetName: "Sheet1",
		Sub: []SubInfo{
			{
				Value1: "value1",
				Value2: "value2",
			},
			{
				Value1: "value3",
				Value2: "value4",
			},
		},
	}
	excelGenerator.RegisterTemplate(SubBlockTemplate, &excelizer.TemplateConfig{
		Name: SubBlockTemplate,
		StaticSheets: []excelizer.SheetConfig{
			{
				Name: "Resp.SheetName",
				Blocks: []excelizer.BlockConfig{
					{
						Type:              excelizer.DynamicBlock,
						SubBlockFieldName: "Resp.Sub",
						SubBlockConfig: &excelizer.BlockConfig{
							Type: excelizer.DynamicCell,
							DynamicCells: []excelizer.CellConfig{
								{Header: "Header1", Value: "Value1"},
								{Header: "Header2", Value: "Value2"},
							},
						},
						ColumnWidths: []float64{50, 50},
					},
				},
			},
		},
	})
	_, _ = excelizer.GetExcelizer().Generate(SubBlockTemplate, excelizer.Param{Resp: resp})
}

const DynamicSheetTemplate string = "DynamicSheetTemplate"

func TestDynamicSheetExcel(t *testing.T) {
	type SubInfo struct {
		Value1 string `json:"value1"`
		Value2 string `json:"value2"`
	}
	resp := []struct {
		SheetName string    `json:"sheetName"`
		Sub       []SubInfo `json:"sub"`
	}{
		{
			SheetName: "Sheet1",
			Sub: []SubInfo{
				{
					Value1: "value1",
					Value2: "value2",
				},
			},
		},
		{
			SheetName: "Sheet2",
			Sub: []SubInfo{
				{
					Value1: "value1",
					Value2: "value2",
				},
			},
		},
	}
	excelGenerator.RegisterTemplate(DynamicSheetTemplate, &excelizer.TemplateConfig{
		Name: DynamicSheetTemplate,
		DynamicSheets: []excelizer.DynamicSheetConfig{
			{
				DynamicFieldName: "Resp",
				SheetNameField:   "SheetName",
				Blocks: []excelizer.BlockConfig{
					{
						Type:              excelizer.DynamicBlock,
						SubBlockFieldName: "Sub",
						SubBlockConfig: &excelizer.BlockConfig{
							Type: excelizer.DynamicCell,
							DynamicCells: []excelizer.CellConfig{
								{Header: "Header1", Value: "Value1"},
								{Header: "Header2", Value: "Value2"},
							},
						},
						ColumnWidths: []float64{50, 50},
					},
				},
			},
		},
	})
	_, _ = excelizer.GetExcelizer().Generate(DynamicSheetTemplate, excelizer.Param{Resp: resp})
}
