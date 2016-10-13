using Microsoft.Office.Interop.Excel;
using System;

namespace ExportToExcelText
{
    class ExcelUtil
    {
        System.Data.DataTable table11 = new System.Data.DataTable();

        public static bool ExportToExcel(System.Data.DataTable table, string saveFileName)
        {

            bool fileSaved = false;

            //ExcelApp xlApp = new ExcelApp();

            Application xlApp = new Application();

            if (xlApp == null)
            {
                return fileSaved;
            }

            Workbooks workbooks = xlApp.Workbooks;
            Workbook workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];//取得sheet1

            long rows = table.Rows.Count;

            /*下边注释的两行代码当数据行数超过行时，出现异常：异常来自HRESULT:0x800A03EC。因为：Excel 2003每个sheet只支持最大行数据

            //Range fchR = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[table.Rows.Count+2, gridview.Columns.View.VisibleColumns.Count+1]];

            //fchR.Value2 = datas;*/

            if (rows > 65535)
            {

                long pageRows = 60000;//定义每页显示的行数,行数必须小于60000

                int scount = (int)(rows / pageRows);

                if (scount * pageRows < table.Rows.Count)//当总行数不被pageRows整除时，经过四舍五入可能页数不准
                {
                    scount = scount + 1;
                }

                for (int sc = 1; sc <= scount; sc++)
                {
                    if (sc > 1)
                    {

                        object missing = System.Reflection.Missing.Value;

                        worksheet = workbook.Worksheets.Add(

                       missing, missing, missing, missing);//添加一个sheet

                    }

                    else
                    {
                        worksheet = (Worksheet)workbook.Worksheets[sc];//取得sheet1
                    }

                    string[,] datas = new string[pageRows + 1, table.Columns.Count + 1];

                    for (int i = 0; i < table.Columns.Count; i++) //写入字段
                    {
                        datas[0, i] = table.Columns[i].Caption;
                    }

                    Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, table.Columns.Count]];
                    range.Interior.ColorIndex = 15;//15代表灰色
                    range.Font.Bold = true;
                    range.Font.Size = 9;

                    int init = int.Parse(((sc - 1) * pageRows).ToString());
                    int r = 0;
                    int index = 0;
                    int result;

                    if (pageRows * sc >= table.Rows.Count)
                    {
                        result = table.Rows.Count;
                    }
                    else
                    {
                        result = int.Parse((pageRows * sc).ToString());
                    }
                    for (r = init; r < result; r++)
                    {
                        index = index + 1;
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            if (table.Columns[i].DataType == typeof(string) || table.Columns[i].DataType == typeof(Decimal) || table.Columns[i].DataType == typeof(DateTime))
                            {
                                object obj = table.Rows[r][table.Columns[i].ColumnName];
                                datas[index, i] = obj == null ? "" : "'" + obj.ToString().Trim();//在obj.ToString()前加单引号是为了防止自动转化格式

                            }

                        }
                    }

                    Range fchR = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[index + 2, table.Columns.Count + 1]];

                    fchR.Value2 = datas;
                    worksheet.Columns.EntireColumn.AutoFit();//列宽自适应。

                    range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[index + 1, table.Columns.Count]];

                    //15代表灰色

                    range.Font.Size = 9;
                    range.RowHeight = 14.25;
                    range.Borders.LineStyle = 1;
                    range.HorizontalAlignment = 1;

                }

            }

            else
            {

                string[,] datas = new string[table.Rows.Count + 2, table.Columns.Count + 1];
                for (int i = 0; i < table.Columns.Count; i++) //写入字段         
                {
                    datas[0, i] = table.Columns[i].Caption;
                }

                Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, table.Columns.Count]];
                range.Interior.ColorIndex = 15;//15代表灰色
                range.Font.Bold = true;
                range.Font.Size = 9;

                int r = 0;
                for (r = 0; r < table.Rows.Count; r++)
                {
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        if (table.Columns[i].DataType == typeof(string) || table.Columns[i].DataType == typeof(Decimal) || table.Columns[i].DataType == typeof(DateTime))
                        {
                            object obj = table.Rows[r][table.Columns[i].ColumnName];
                            datas[r + 1, i] = obj == null ? "" : "'" + obj.ToString().Trim();//在obj.ToString()前加单引号是为了防止自动转化格式

                        }

                    }

                    //System.Windows.Forms.Application.DoEvents();

                }

                Range fchR = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[table.Rows.Count + 2, table.Columns.Count + 1]];

                fchR.Value2 = datas;

                worksheet.Columns.EntireColumn.AutoFit();//列宽自适应。

                range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[table.Rows.Count + 1, table.Columns.Count]];

                //15代表灰色

                range.Font.Size = 9;
                range.RowHeight = 14.25;
                range.Borders.LineStyle = 1;
                range.HorizontalAlignment = 1;
            }

            if (saveFileName != "")
            {
                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs(saveFileName);
                    fileSaved = true;

                }

                catch (Exception ex)
                {
                    fileSaved = false;
                }

            }

            else
            {

                fileSaved = false;

            }

            xlApp.Quit();

            GC.Collect();//强行销毁 
            return fileSaved;
        }
    }
}
