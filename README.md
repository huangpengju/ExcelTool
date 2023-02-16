# ——————————firstExcelTool包——————————
## 1.数字教材导入，生成内外部标签表和一系列文件夹
```go
firstexceltool.CreateTwoExcel()
// 读取目标Excel，创建新的Excel和一系列文件夹<br>
// 在vcom工作时，为了提升工作效率而编写，适用场景是数字教材入库前，对表格数据的批量处理 2023/2/7<br>
// 学习并使用了以下主要内容：
// os库和io/ioutil库
```
```go
str, _ := os.Getwd()//获取当前路径

infos, err := ioutil.ReadDir(str)//在路径str中读取全部文件

os.Mkdir()//创建文件夹
```
### strings库
```go
func strings.LastIndex(s string, substr string) int//获取substr在s中的位置，未找到返回-1
```
### excelize库<br>
Excelize 是 Go 语言编写的用于操作 Office Excel 文档基础库
https://xuri.me/excelize/zh-hans/
```go
excelize.OpenFile(filename string) //打开指定的工作簿
GetSheetName(1)  //获取第1个Sheet表名
GetRows(name)   //读取name表数据      
```
```go
excelize.NewFile()//创建一个工作簿
NewSheet("Sheet1")//创建Sheet1工作表
SetCellValue("Sheet1", "A1", "ID")//设置Sheet1中A1单元格的值
SetColWidth("Sheet1", "A", "B", 30)//设置Sheet1表第A列到B列的列宽
SaveAs("内部标签表.xlsx")//把工作簿保存并命名为“内部标签表.xlsx”
```
# ——————————secondExcelTool包——————————
## 2.普通资源导入,生成目录路径
```go
secondexceltool.CreateOneExcel()
// 读取目标Excel，创建新的Excel<br>
// 在vcom工作时，为了提升工作效率而编写，适用场景是备授课资源入库前，一键生成入库资源单子 2023/2/13<br>
// 学习并使用了以下主要内容：
// os库和io/ioutil库
```
```go
str, _ := os.Getwd()//获取当前路径

infos, err := ioutil.ReadDir(str)//在路径str中读取全部文件

```
### strings包
```go
func strings.LastIndex(s string, substr string) int//获取substr在s中的位置，未找到返回-1
strings.HasSuffix()  ////判断参数2，是不是在参数1中
strings.TrimSuffix()  //// 返回去除s可能的后缀suffix的字符串。
```
### excelize库<br>
Excelize 是 Go 语言编写的用于操作 Office Excel 文档基础库
https://xuri.me/excelize/zh-hans/
```go
excelize.OpenFile(filename string) //打开指定的工作簿
GetSheetName(1)  //获取第1个Sheet表名
GetRows(sheetName1)   //读取sheetName1表数据 


excelize.NewFile()  //创建一个工作薄
NewSheet("Sheet1")// 创建一个工作表
SetColWidth("Sheet1", "A", "B", 80)//设置列宽
SetCellValue("Sheet1", "A1", "基础目录")	//设置单元格内容
SetActiveSheet(index)// 设置工作簿的默认工作表
SaveAs("基础目录合成.xlsx")//另存为工作薄“基础目录合成.xlsx”
```
### builtin包<br>
```go
append()   //内建函数append将元素追加到切片的末尾。
```
