# VantageOffice

VantageOffice 是一套纯 Go 语言实现的办公文档处理库，涵盖 Word、Excel、PowerPoint 三大办公文档格式。所有组件零外部依赖，API 设计简洁，适合在服务端、CLI 工具或任何 Go 项目中生成和解析办公文档。

## 套件总览

| 组件 | 格式 | 仓库 | 测试覆盖率 |
|------|------|------|-----------|
| **GoWord** | `.docx` (Office Open XML) | [VantageDataChat/GoWord](https://github.com/VantageDataChat/GoWord) | 93.3%–100% |
| **GoExcel** | `.xlsx` / `.csv` | [VantageDataChat/GoExcel](https://github.com/VantageDataChat/GoExcel) | 97.7% |
| **GoPPT** | `.pptx` (OOXML) | [VantageDataChat/GoPPT](https://github.com/VantageDataChat/GoPPT) | 96% |

## 核心特点

- **纯 Go 实现** — 无 CGO、无外部依赖，交叉编译友好
- **读写双向** — 三个组件均支持创建新文档和读取已有文档
- **高测试覆盖** — 全部组件覆盖率 93%+，可放心用于生产环境
- **MIT 许可** — 商业友好，可自由使用

---

## GoWord — Word 文档处理

纯 Go 的 `.docx` 读写库，灵感来自 PHPWord。

### 安装

```bash
go get github.com/VantageDataChat/GoWord
```

### 功能亮点

- 文档属性（标题、作者、主题、关键词等）
- 分节与页面布局（纸张大小、方向、页边距、分栏、页码）
- 段落与富文本（粗体、斜体、下划线、颜色、字号等）
- 标题（Title、Heading1–9）
- 表格（单元格合并、边框、底纹、嵌套表格）
- 列表（项目符号、数字编号、罗马数字等）
- 图片（PNG、JPEG、GIF、BMP、TIFF，支持文件路径或字节数据）
- 页眉页脚与页码
- 脚注、尾注、批注、书签、目录
- 复选框、线条形状
- 命名样式（字体、段落、表格、编号）
- TextRun 混合格式内联内容
- 单位转换工具（twip、cm、inch、pt、EMU、pixel）

### 快速示例

```go
package main

import (
    "log"
    goword "github.com/VantageDataChat/GoWord"
    "github.com/VantageDataChat/GoWord/style"
)

func main() {
    doc := goword.New()
    doc.Properties.Title = "My Document"

    sec := doc.AddSection()
    sec.AddTitle("Hello GoWord", 1)
    sec.AddText("This is bold text.",
        &style.FontStyle{Bold: true, Size: 12, Color: "333333"}, nil)

    if err := doc.Save("hello.docx"); err != nil {
        log.Fatal(err)
    }
}
```

---

## GoExcel — Excel / CSV 处理

纯 Go 的电子表格处理库，灵感来自 PHPOffice/PhpSpreadsheet。支持 XLSX 和 CSV 格式。

### 安装

```bash
go get github.com/VantageDataChat/GoExcel
```

### 功能亮点

- XLSX 读写（兼容 Excel 2007+）
- CSV 读写（自定义分隔符）
- 公式计算引擎（24 个内置函数：SUM、AVERAGE、IF、COUNTIF 等）
- 完整样式系统（字体、边框、填充、对齐、数字格式）
- 合并单元格、冻结窗格
- 行列插入 / 删除 / 复制
- 超链接、批注、富文本
- 条件格式、数据验证、自动筛选
- 页面设置与打印配置
- 工作表 / 工作簿保护
- 文档属性、命名范围
- Unicode 和特殊字符支持

### 支持的公式函数

| 类别 | 函数 |
|------|------|
| 数学 | SUM, ABS, ROUND, SQRT, POWER, MOD, INT |
| 统计 | AVERAGE, COUNT, COUNTA, MAX, MIN, MEDIAN |
| 逻辑 | IF |
| 文本 | LEN, UPPER, LOWER, TRIM, LEFT, RIGHT, MID, CONCATENATE |
| 条件 | SUMIF, COUNTIF |

### 快速示例

```go
package main

import "github.com/VantageDataChat/GoExcel"

func main() {
    wb := gospreadsheet.New()
    ws := wb.GetActiveSheet()

    ws.SetCellValue("A1", "姓名")
    ws.SetCellValue("B1", "分数")
    ws.SetCellValue("A2", "Alice")
    ws.SetCellValue("B2", 95.5)

    ws.SetCellFormula("B3", "AVERAGE(B1:B2)")

    ws.SetCellStyle("A1", gospreadsheet.NewStyle().
        SetFont(&gospreadsheet.Font{Bold: true, Size: 14}))

    gospreadsheet.SaveFile(wb, "output.xlsx")
}
```

---

## GoPPT — PowerPoint 演示文稿处理

纯 Go 的 `.pptx` 创建、读取和写入库，灵感来自 PHPOffice/PHPPresentation。

### 安装

```bash
go get github.com/VantageDataChat/GoPPT
```

### 功能亮点

- 创建和保存 `.pptx` 文件（PowerPoint 2007+）
- 读取现有 `.pptx` 文件，支持完整读写往返
- 富文本（字体、颜色、粗体、斜体、下划线、删除线）
- 图片（PNG、JPEG、GIF、BMP、SVG）
- 表格（单元格格式和填充）
- 自动形状（矩形、椭圆、三角形、箭头、星形等）
- 线条形状
- 图表：柱状图、3D 柱状图、折线图、面积图、饼图、3D 饼图、环形图、散点图、雷达图
- 组合形状、占位符形状
- 项目符号（字符和数字编号）
- 批注（含作者信息）、演讲者备注
- 幻灯片背景（纯色和渐变）
- 动画（基础分组）
- 文档属性和自定义属性
- 多种幻灯片布局（4:3、16:9、16:10、A4、Letter、自定义）

### 快速示例

```go
package main

import (
    "log"
    ppt "github.com/VantageDataChat/GoPPT"
)

func main() {
    p := ppt.New()
    p.GetDocumentProperties().Title = "My Presentation"

    slide := p.GetActiveSlide()

    title := slide.CreateRichTextShape()
    title.SetOffsetX(500000).SetOffsetY(300000)
    title.SetWidth(8000000).SetHeight(1000000)
    tr := title.CreateTextRun("Hello, GoPPT!")
    tr.GetFont().SetSize(28).SetBold(true).SetColor(ppt.ColorBlue)

    slide2 := p.CreateSlide()
    chart := slide2.CreateChartShape()
    chart.BaseShape.SetOffsetX(500000).SetOffsetY(500000)
    chart.BaseShape.SetWidth(7000000).SetHeight(4500000)
    chart.GetTitle().SetText("Sales Report")

    bar := ppt.NewBarChart()
    bar.AddSeries(ppt.NewChartSeriesOrdered("Revenue",
        []string{"Q1", "Q2", "Q3", "Q4"},
        []float64{120, 180, 150, 210},
    ))
    chart.GetPlotArea().SetType(bar)

    w, _ := ppt.NewWriter(p, ppt.WriterPowerPoint2007)
    if err := w.(*ppt.PPTXWriter).Save("presentation.pptx"); err != nil {
        log.Fatal(err)
    }
}
```

---

## 许可证

VantageOffice 套件中的所有组件均采用 [MIT License](https://opensource.org/licenses/MIT) 发布。
