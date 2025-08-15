package test

import (
	"excelizer"
	"testing"
	"time"
)

type TimeRange struct {
	StartTime time.Time `json:"startTime"`
	EndTime   time.Time `json:"endTime"`
}

func TestStaticExcel(t *testing.T) {
	Req := TimeRange{
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
	excelGenerator := excelizer.GetExcelizer()
	excelGenerator.RegisterTemplate(StaticTemplate, CreateStaticTemplate())
	_, _ = excelGenerator.Generate(StaticTemplate, excelizer.Param{Req: Req, Resp: resp})
}

func TestSubBlockExcel(t *testing.T) {
	Req := TimeRange{
		StartTime: time.Now(),
		EndTime:   time.Now(),
	}
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
	excelGenerator := excelizer.GetExcelizer()
	excelGenerator.RegisterTemplate(SubBlockTemplate, CreateSubBlockTemplate())
	_, _ = excelizer.GetExcelizer().Generate(SubBlockTemplate, excelizer.Param{Req: Req, Resp: resp})
}

func TestDynamicSheetExcel(t *testing.T) {
	Req := TimeRange{
		StartTime: time.Now(),
		EndTime:   time.Now(),
	}
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
	excelGenerator := excelizer.GetExcelizer()
	excelGenerator.RegisterTemplate(DynamicSheetTemplate, CreateDynamicSheetTemplate())
	_, _ = excelizer.GetExcelizer().Generate(DynamicSheetTemplate, excelizer.Param{Req: Req, Resp: resp})
}
