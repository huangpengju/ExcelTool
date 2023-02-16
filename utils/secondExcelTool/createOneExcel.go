package secondexceltool

import (
	"fmt"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

func CreateOneExcel() {
	// CreateExcel()
	ReadExcel()
}

// 读取Excel文档
func ReadExcel() {
	//1.GetFilePath() 获取"(教材目录).xlsx"文件-全名
	fileName := GetFilePath()
	if fileName == "-1" {
		return
	}
	//2.打开fileName文件
	f, err := excelize.OpenFile(fileName)

	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// sheetName := f.GetSheetName(0)
	sheetName1 := f.GetSheetName(1) //获取第2个Sheet表名
	// println(sheetName1)

	// 获取 SheetName 上所有单元格
	rows, err := f.GetRows(sheetName1)
	if err != nil {
		fmt.Println(err)
		return
	}
	// fmt.Println(rows)
	rowsDir := DeleteSlice(rows, 0)    //把表中的第一行的值删除，并返回表中的其他值
	sourceName := rowsSlice(rowsDir)   //处理路径切片，返回课名切片
	getSourcePath(sourceName, rowsDir) //根据课名切片，创建表
}

// 获取"(教材目录).xlsx"文件-全名
func GetFilePath() (fileName string) {
	fileName = "-1" //初始化文件名
	//获取当前工作目录（路径）
	filePath, _ := os.Getwd()
	//读取工作目录
	files, _ := os.ReadDir(filePath)

	for _, file := range files {
		if !file.IsDir() {
			if strings.HasSuffix(file.Name(), "(教材目录).xlsx") { //判断参数2，是不是在参数1中
				return file.Name()
			}
		}

	}
	return fileName
}

// 创建Excel文档
func CreateExcel(sourceDir []string) {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// 创建一个工作表
	index, err := f.NewSheet("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	//设置列宽
	err = f.SetColWidth("Sheet1", "A", "B", 250)
	if err != nil {
		fmt.Println(err)
		return
	}
	//设置标题
	err = f.SetCellValue("Sheet1", "A1", "基础目录")
	if err != nil {
		fmt.Println(err)
		return
	}

	// 设置单元格的值
	var rowNum = 1 //行号
	for _, val := range sourceDir {
		rowNum++
		// 设置单元格的值
		f.SetCellValue("Sheet1", fmt.Sprintf("%v%d", "A", rowNum), val)
	}

	// 设置工作簿的默认工作表
	f.SetActiveSheet(index)
	// 根据指定路径保存文件
	if err := f.SaveAs("基础目录合成.xlsx"); err != nil {
		fmt.Println(err)
	}

}

// 根据课名，合并路径
func getAddDir(sourceDir, sourceName []string) (dirName []string) {
	var name string
	for _, val := range sourceName {
		// fmt.Println(val)
		for _, dir := range sourceDir {
			if strings.LastIndex(dir, val) != -1 {
				name += dir + "\r\n" //给原表中每行数据添加 换行和回车符
				// fmt.Println(name)
			}
		}
		//删除最后一个位置多余的 换行和回车符
		dirName = append(dirName, strings.TrimSuffix(name, "\r\n")) //func TrimSuffix(s, suffix string) string   返回去除s可能的后缀suffix的字符串。
		// fmt.Println(name)

		name = ""

	}
	return dirName
}

// 根据课程名，筛选路径
func getSourcePath(sourceName []string, rowsDir []string) {
	var sourceDir []string
	for _, row := range rowsDir {
		for _, val := range sourceName {
			if strings.LastIndex(row, val) != -1 {
				sourceDir = append(sourceDir, row)
			}
		}

	}
	dir := getAddDir(sourceDir, sourceName)
	//创建表格
	CreateExcel(dir)
}

// 处理切片,获取课的名称
func rowsSlice(rowsDir []string) (sourceName []string) {
	var name string
	for _, colCell := range rowsDir {

		if strings.LastIndex(colCell, "/课时") != -1 {
			str1 := colCell[:strings.LastIndex(colCell, "/课时")]
			if name != colCell[strings.LastIndex(str1, "/")+1:strings.LastIndex(colCell, "/课时")] {
				name = colCell[strings.LastIndex(str1, "/")+1 : strings.LastIndex(colCell, "/课时")]
				sourceName = append(sourceName, name)
			}

		}
	}
	return sourceName
}

// 删除切片中的第1个元素。
func DeleteSlice(a [][]string, elem int) (t1 []string) {
	tmp1 := make([]string, 0, len(a)-1) //定义一个空切片
	for i, v := range a {
		if i != elem {
			tmp1 = append(tmp1, v[0]) //返回表单元格内容
		}
	}
	return tmp1
}
