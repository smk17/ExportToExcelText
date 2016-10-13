using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportToExcelText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void 导出EToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            //设置文件类型
            saveFileDialog.Filter = "xlsx files(*.xlsx)|*.xlsx|xls files(*.xls)|*.xls|All files(*.*)|*.*";
            //设置默认文件名（可以不设置）
            saveFileDialog.FileName = "ExportToExcelText-" + DateTime.Now.ToString("yyyyMMdd");
            //主设置默认文件extension（可以不设置）
            saveFileDialog.DefaultExt = "xlsx";
            //获取或设置一个值，该值指示如果用户省略扩展名，文件对话框是否自动在文件名中添加扩展名。（可以不设置）
            saveFileDialog.AddExtension = true;
            //保存对话框是否记忆上次打开的目录
            saveFileDialog.RestoreDirectory = true;
            // Show save file dialog box
            DialogResult result = saveFileDialog.ShowDialog();
            //点了保存按钮进入
            if (result == DialogResult.OK)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Age", typeof(string));
                for (int i = 0; i < 10; i++)
                {
                    DataRow dr = dt.NewRow();
                    dr["Name"] = "spring" + i.ToString();
                    dr["Age"] = "2" + i.ToString();
                    dt.Rows.Add(dr);
                }
                dt.AcceptChanges();

                ExcelUtil.ExportToExcel(dt, saveFileDialog.FileName.ToString());
            }
        }
    }
}
