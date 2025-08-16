package test

import (
	"excelizer"
	"log"
	"testing"
	"time"
)

func TestSimpleExcel(t *testing.T) {
	var excelGenerator = excelizer.GetExcelizer()

	Req := struct {
		StartTime time.Time `json:"startTime"`
		EndTime   time.Time `json:"endTime"`
	}{
		StartTime: time.Time{},
		EndTime:   time.Now(),
	}

	const templateName = "simpleTemplate"
	template := &excelizer.TemplateConfig{
		Name: templateName,
		StaticSheets: []excelizer.SheetConfig{
			{
				Name: "simple",
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
				},
			},
		},
	}
	excelGenerator.RegisterTemplate(templateName, template)

	filepath, err := excelGenerator.Generate(templateName, excelizer.Param{Req: Req})
	if err != nil {
		log.Fatal(err)
	}
	log.Printf("excel generated at: %v", filepath)
}
