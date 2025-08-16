# Excelizer

将已有结果，通过定义模板，导出为excel文件的工具库。本项目基于[excelize](https://github.com/qax-os/excelize)开发。

## Why Excelizer？

1. 将已有结果导出excel时，手动指定单元格位置（如 A1, B2）既麻烦又容易出错，实现过程可以说全是硬编码，后期维护可谓牵一发而动全身
2. 对于可变行数，可变sheet页的处理困难

Excelizer 通过给已有结果定义一个模板，自动生成Excel，支持动态生成可变行数、可变表格块、可变sheet页的复杂数据。另有内置的数据转换，单元格合并，列宽配置，单元格样式配置，图表生成等功能。更多功能可以在此框架基础上二次开发。

## How it works?
Excelizer 通过定义模板配置来描述 Excel 文件的结构和数据映射关系。模板包含静态和动态 Sheet 配置，每个 Sheet 包含若干个数据块(Block)，每个块又包含若干个单元格，支持静态单元格、动态单元格和动态块三种类型。

核心概念：
- TemplateConfig: 定义整个 Excel 文件的模板配置
- SheetConfig: 定义静态 Sheet 的结构和内容
- DynamicSheetConfig: 定义动态 Sheet，根据数据动态生成多个 Sheet
- BlockConfig: 定义数据块，支持静态单元格、动态单元格和动态块三种类型
- CellConfig: 定义单元格配置，包括值、样式、格式化等

使用时只需定义好模板，提供数据，Excelizer 会自动渲染生成完整的 Excel 文件，无需手动操作单元格位置。

## Quick Start
只需三步：待导出的结果 -> 定义对应模板并注册 -> 自动生成excel文件

来看一个最简实现（test/quickStart_test中TestSimpleExcel）：
1. 待导出的结果
```go
Req := struct {
    StartTime time.Time `json:"startTime"`
    EndTime   time.Time `json:"endTime"`
}{
    StartTime: time.Time{},
    EndTime:   time.Now(),
}
```

2. 定义对应模板并注册
```go
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
```

3. 自动生成excel文件
```go
_, _ := excelGenerator.Generate(templateName, excelizer.Param{Req: Req})
```

导出内容预览：
![](<test/img/Pasted image 20250816113559.png>)

## Feature
以下导出结果对应test/template_test下的测试函数，如：
一张sheet页中同时包含静态块和动态块：
![alt text](<test/img/Pasted image 20250815165920.png>)

一张sheet页中包含多个动态块，中间间隔多个空行
![alt text](<test/img/Pasted image 20250815165949.png>)

可变sheet页：
![alt text](<test/img/Pasted image 20250815170039.png>)

