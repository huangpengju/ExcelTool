package firstexceltool

import (
	"fmt"
	"io/ioutil"
	"os"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func CreateTwoExcel() {
	//1.获取当前目录下的XXXX(教材目录).xlsx文件名
	info, _ := fileName()
	//2.判断是否获取到文件全名
	if info != "-1" {
		//3.数据处理，打开找到的XXXXXXX(教材目录).xlsx文件（工作簿）
		OpenExcel(info)
		println("已生成表格！！！")
	} else {
		println("无表格需要处理！！！")
	}
}

// 获取当前目录下的表格文件名
func fileName() (fileName string, err error) {
	str, _ := os.Getwd()

	infos, err := ioutil.ReadDir(str) //读取全部文件
	if err != nil {
		return "读取失败", err
	}

	fileName = "-1" //工作簿名初始化
	//遍历文件名
	for _, info := range infos {
		// comma := strings.LastIndex(info.Name(), ".") //获取最后一个点的位置

		// if fileType == info.Name()[comma:] { //通过后缀名判断是不是要找的同类型文件
		if strings.HasSuffix(info.Name(), "(教材目录).xlsx") {

			comma1 := strings.LastIndex(info.Name(), "(教材目录)") //判断表名中是否含有某些特征
			if comma1 != -1 {                                  //通过文件最后有没有字符“(教材目录)”，判断是不是导出单子
				// fmt.Println(info.Name()) //获取文件的名称
				fileName = info.Name() //工作簿名再赋值
			}
		}

	}

	return fileName, err
}

// 打开指定excel工作簿
func OpenExcel(info string) {
	//4.打开excel工作簿
	f, err := excelize.OpenFile(info)
	if err != nil {
		fmt.Println(err)
		return
	}
	name := f.GetSheetName(1)       //获取第1个Sheet表名
	rows := f.GetRows(name)         //读取第1个Sheet表数据
	t1, t2 := DeleteSlice2(rows, 0) //删除Sheet表中的第一行标题,返回第1个Sheet表中所有行的id和目录路径

	//5.创建excel工作簿
	CreateExcel(t1, t2)

}

// 创建指定excel工作簿
func CreateExcel(t1, t2 []string) {
	//6.创建一个excel工作簿(内部标签表)
	f1 := excelize.NewFile()
	//6.1创建一个工作表
	sheetNum1 := f1.NewSheet("Sheet1")

	//7.创建一个excel工作簿（外部标签表）
	f2 := excelize.NewFile()
	//7.1创建一个工作表
	sheetNum2 := f2.NewSheet("Sheet1")

	//开始设置单元格的值

	//6.2设置内部标签第1行的表头标题
	f1.SetCellValue("Sheet1", "A1", "ID")
	f1.SetCellValue("Sheet1", "B1", "课程名")
	//6.2设置内部标签第1行的列宽
	f1.SetColWidth("Sheet1", "A", "B", 30)
	f1.SetColWidth("Sheet1", "B", "C", 100)

	//7.2设置外部标签第1行的表头标题
	f2.SetCellValue("Sheet1", "A1", "教材目录ID")
	f2.SetCellValue("Sheet1", "B1", "教材目录")
	f2.SetCellValue("Sheet1", "C1", "同步培优目录ID")
	f2.SetCellValue("Sheet1", "D1", "同步培优目录")
	f2.SetCellValue("Sheet1", "E1", "专题目录ID")
	f2.SetCellValue("Sheet1", "F1", "专题目录")
	f2.SetCellValue("Sheet1", "G1", "ID")
	f2.SetCellValue("Sheet1", "H1", "图片")
	f2.SetCellValue("Sheet1", "I1", "是否有热区")
	f2.SetCellValue("Sheet1", "J1", "单词跳转目录")
	f2.SetCellValue("Sheet1", "K1", "displayname")
	//7.2设置内部标签第1行的列宽
	f2.SetColWidth("Sheet1", "A", "B", 30)
	f2.SetColWidth("Sheet1", "B", "C", 100)
	f2.SetColWidth("Sheet1", "C", "F", 10)
	f2.SetColWidth("Sheet1", "G", "H", 30)
	f2.SetColWidth("Sheet1", "H", "J", 10)
	f2.SetColWidth("Sheet1", "K", "L", 50)

	timeUnix := time.Now().Unix() //当前时间戳单位s的时间戳
	timeEnd := "2023-10-07"
	//使用Parse 默认获取为UTC时区 需要获取本地时区 所以使用ParseInLocation
	time1, _ := time.ParseInLocation("2006-01-02", timeEnd, time.Local)
	towTimeStamp := time1.AddDate(0, 0, 1).Unix()
	// fmt.Println("明天：", towTimeStamp) // 2022-05-01
	if timeUnix < towTimeStamp {
		//8.给表格的第2行写入数据
		rowNum := 1 //初始化行号（第2行的下标）

		//遍历读目录路径（此处代码可以优化，应该可以用递归处理）
		for i, row := range t2 {
			comma := strings.LastIndex(row, "课时") //获取最后“课时”的位置
			if comma != -1 {                      //如果“课时”存在，那么课时这一行数据的上一层，可能就是课（课时在课的下层）
				comma1 := strings.LastIndex(t2[i-1], "课时") //再次判断数据中是否含有“课时”,i-1就是上一层级，还要判断i-1这一层级没有课时
				if comma1 == -1 {                          //没有“课时了”
					rowNum++ //让行号自增

					f1.SetCellValue("Sheet1", fmt.Sprintf("%v%d", "A", rowNum), t1[i-1]) //给内部标签表 A列  写入id
					//处理资源名	// println("课程名", t2[i-1][comma2+1:])，校本资源：《第10课 音频处理我最棒》数字教材
					comma2 := strings.LastIndex(t2[i-1], "/")
					bookName := fmt.Sprintf("%v%v%v", "校本资源：《", t2[i-1][comma2+1:], "》数字教材")
					f1.SetCellValue("Sheet1", fmt.Sprintf("%v%d", "B", rowNum), bookName) ////给内部标签表 B列  写入资源名

					f2.SetCellValue("Sheet1", fmt.Sprintf("%v%d", "A", rowNum), t1[i-1])  //给外部标签表A 列   写入id
					f2.SetCellValue("Sheet1", fmt.Sprintf("%v%d", "B", rowNum), t2[i-1])  //给外部标签表B 列   写入目录
					f2.SetCellValue("Sheet1", fmt.Sprintf("%v%d", "G", rowNum), t1[i-1])  //给外部标签表G 列   写入 id
					f2.SetCellValue("Sheet1", fmt.Sprintf("%v%d", "I", rowNum), 0)        //给外部标签表I 列   写入 默认值0
					f2.SetCellValue("Sheet1", fmt.Sprintf("%v%d", "K", rowNum), bookName) //给外部标签表K 列   写入资源名

					CreateDir(t1[i-1]) //创建文件夹
				}
			}
		}
	} else {
		f1.SetCellValue("Sheet1", "A2", "数据处理过期")
		f2.SetCellValue("Sheet1", "A2", "数据处理过期")

	}
	//设置工作簿的默认工作表
	f1.SetActiveSheet(sheetNum1)
	f1.SetActiveSheet(sheetNum2)
	//根据指定路径保存文件
	if err := f1.SaveAs("内部标签表.xlsx"); err != nil {
		fmt.Println(err)
	}
	//根据指定路径保存文件
	if err := f2.SaveAs("外部标签表.xlsx"); err != nil {
		fmt.Println(err)
	}

}

// 删除切片中的第1个元素。
func DeleteSlice2(a [][]string, elem int) (t1, t2 []string) {
	tmp1 := make([]string, 0, len(a)-1)
	tmp2 := make([]string, 0, len(a)-1)
	for i, v := range a {
		if i != elem {
			tmp1 = append(tmp1, v[0])
			tmp2 = append(tmp2, v[1])
		}
	}
	return tmp1, tmp2
}

// 判断文件夹是否存在
func HasDir(path string) (bool, error) {
	_, _err := os.Stat(path)
	if _err == nil {
		return true, nil
	}
	if os.IsNotExist(_err) {
		return false, nil
	}
	return false, _err
}

// 创建文件夹
func CreateDir(path string) {
	_exist, _err := HasDir(path)
	if _err != nil {
		fmt.Printf("获取文件夹异常 -> %v\n", _err)
		return
	}
	if _exist {
		fmt.Println("文件夹已存在！")
	} else {
		err := os.Mkdir(path, os.ModePerm)
		if err != nil {
			fmt.Printf("创建文件夹目录异常 -> %v\n", err)
		} else {
			fmt.Println("文件夹创建成功!")
		}
	}
}
