package main

import (
	firstexceltool "ExcelTool/utils/firstExcelTool"
	secondexceltool "ExcelTool/utils/secondExcelTool"
)

func main() {
	//1.数字教材导入，生成内外部标签表和一系列文件夹
	firstexceltool.CreateTwoExcel()
	//2.普通资源导入,生成目录路径
	secondexceltool.CreateOneExcel()
}
