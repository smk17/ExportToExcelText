# ExportToExcelText
DoNet下大数据量导出到Excel
  
# 添加引用
添加Microsoft.Office.Interop.Excel引用，在我得示例DEMO得根目录下就有Microsoft.Office.Interop.Excel.dll 这个文件，下载我的DEMO复制这个DLL文件到你的项目并添加这个DLL到你得项目引用即可。

# Excel类的简单介绍 
此命名空间下关于Excel类的结构分别为： 
> * ApplicationClass - 就是我们的excel应用程序。 
> * Workbook - 就是我们平常见的一个个excel文件，经常是使用Workbooks类对其进行操作。 
> * Worksheet - 就是excel文件中的一个个sheet页。 
Worksheet.Cells[row, column] - 就是某行某列的单元格，注意这里的下标row和column都是从1开始的，跟我平常用的数组或集合的下标有所不同。 
知道了上述基本知识后，利用此类来操作excel就清晰了很多。
