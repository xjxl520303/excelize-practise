package main

import (
	"log"
	"strconv"

	"github.com/xuri/excelize/v2"
)

var (
	titleStyle, cellCenterStyle, CellVACStyle,
	cellVACAndVALStyle int
	sheetName  = "Sheet1" // 工作表名称
	titleColor = "92d050" // 标题颜色
	markMsg    = "补考"     // 标记信息
)

func main() {
	f, err := excelize.OpenFile("source.xlsx")
	if err != nil {
		log.Fatal(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	// 删除第11行
	if err := f.RemoveRow(sheetName, 11); err != nil {
		log.Fatal(err)
		return
	}

	// 合并单元格 A1:F1
	if err := f.MergeCell(sheetName, "A1", "F1"); err != nil {
		log.Fatal(err)
		return
	}

	// 创建标题样式
	if titleStyle, err = f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Family:    "黑体",     // 字体
			Size:      18,       // 字号
			Color:     "FFFFFF", // 字体颜色
			Underline: "single", // 下划线，可选值为 none, single，double
		},
		// 水平居中
		Alignment: &excelize.Alignment{Horizontal: "center"},
		// 填充背景色, Pattern 其它值参考文档：样式/图案填充部分
		Fill: excelize.Fill{Type: "pattern", Color: []string{titleColor}, Pattern: 1},
	}); err != nil {
		log.Fatal(err)
		return
	}

	// 创建居中样式
	if cellCenterStyle, err = f.NewStyle(&excelize.Style{
		// 水平居中
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
	}); err != nil {
		log.Fatal(err)
		return
	}

	// 创建垂直居中样式
	if CellVACStyle, err = f.NewStyle(&excelize.Style{
		// 水平居中
		Alignment: &excelize.Alignment{Vertical: "center"},
	}); err != nil {
		log.Fatal(err)
		return
	}

	// 创建垂直居中，水平向左对齐样式
	if cellVACAndVALStyle, err = f.NewStyle(&excelize.Style{
		// 水平居中
		Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"},
	}); err != nil {
		log.Fatal(err)
		return
	}

	// 设置标题行样式
	if err := f.SetCellStyle(sheetName, "A1", "F1", titleStyle); err != nil {
		log.Fatal(err)
		return
	}

	var rowHeight float64
	if rowHeight, err = f.GetRowHeight(sheetName, 1); err != nil {
		log.Fatal(err)
		return
	}

	// 调整行高
	newHeight := rowHeight + float64(12)
	if err := f.SetRowHeight(sheetName, 1, newHeight); err != nil {
		log.Fatal(err)
		return
	}

	// 居中【姓名】列
	if err = f.SetCellStyle(sheetName, "A2", "A18", cellCenterStyle); err != nil {
		log.Fatal(err)
		return
	}

	// 居中第二行
	if err = f.SetCellStyle(sheetName, "A2", "F2", cellCenterStyle); err != nil {
		log.Fatal(err)
		return
	}

	// 设置 B3:E18 样式
	if err = f.SetCellStyle(sheetName, "B3", "E18", CellVACStyle); err != nil {
		log.Fatal(err)
		return
	}

	// 设置 B3:E18 样式
	if err = f.SetCellStyle(sheetName, "F3", "F18", cellVACAndVALStyle); err != nil {
		log.Fatal(err)
		return
	}

	rows, err := f.GetRows(sheetName)

	if err != nil {
		log.Fatal(err)
		return
	}

	for index, row := range rows[2:] {
		var totalScore float64 = 0.0
		for i, val := range row[1:4] {
			score, err := strconv.ParseFloat(val, 32)

			if err != nil {
				log.Fatal(err)
				continue
			}

			if i < len(row[1:4])-1 {
				totalScore += float64(0.3) * score
			} else {
				totalScore += float64(0.4) * score
			}

		}

		// 写入【总成绩】列
		if err := f.SetCellValue(sheetName, "E"+strconv.Itoa(index+3), totalScore); err != nil {
			log.Fatal(err)
			continue
		}

		if totalScore < 60 {
			// 写入【补考否】列
			if err := f.SetCellValue(sheetName, "F"+strconv.Itoa(index+3), markMsg); err != nil {
				log.Fatal(err)
				continue
			}
		}

		totalScore = 0.0
	}

	// 另存为别的文件，方便查看
	if err := f.SaveAs("result.xlsx"); err != nil {
		log.Fatal(err)
		return
	}
}
