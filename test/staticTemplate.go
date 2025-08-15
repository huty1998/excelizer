package test

import (
	"excelizer"
)

const StaticTemplate string = "StaticTemplate"

func CreateStaticTemplate() *excelizer.TemplateConfig {
	return &excelizer.TemplateConfig{
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
	}
}
