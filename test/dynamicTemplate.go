package test

import "excelizer"

const SubBlockTemplate string = "SubBlockTemplate"

func CreateSubBlockTemplate() *excelizer.TemplateConfig {
	return &excelizer.TemplateConfig{
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
	}
}

const DynamicSheetTemplate string = "DynamicSheetTemplate"

func CreateDynamicSheetTemplate() *excelizer.TemplateConfig {
	return &excelizer.TemplateConfig{
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
	}
}
