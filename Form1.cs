using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSplitter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // 拆分按钮点击事件
        private void btnSplit_Click(object sender, EventArgs e)
        {
            string sourceFile = txtFilePath.Text;
            string outputFolder = Path.Combine(Path.GetDirectoryName(sourceFile), Path.GetFileNameWithoutExtension(sourceFile));

            if (!File.Exists(sourceFile))
            {
                MessageBox.Show("源文件不存在！");
                return;
            }

            lblProgress.Text = $"正在拆分...";
            Application.DoEvents();

            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook sourceWorkbook = null;

            try
            {
                if (Path.GetExtension(sourceFile).ToLower() == ".xls")
                {
                    sourceWorkbook = excelApp.Workbooks.Open(sourceFile, 0, false, 5);
                }
                else if (Path.GetExtension(sourceFile).ToLower() == ".xlsx")
                {
                    sourceWorkbook = excelApp.Workbooks.Open(sourceFile, 0, false, 1);
                }
                else
                {
                    MessageBox.Show("不支持的文件格式！");
                    return;
                }

                excelApp.DisplayAlerts = false;
                int sheetCount = sourceWorkbook.Worksheets.Count;
                progressBar.Maximum = sheetCount;

                for (int i = 1; i <= sheetCount; i++)
                {
                    Excel.Worksheet sheet = sourceWorkbook.Worksheets[i];
                    string sheetName = sheet.Name;
                    string outputPath = Path.Combine(outputFolder, $"{sheetName}{Path.GetExtension(sourceFile)}");

                    Excel.Workbook newWorkbook = excelApp.Workbooks.Add();

                    while (newWorkbook.Worksheets.Count > 1)
                    {
                        newWorkbook.Worksheets[1].Delete();
                    }

                    sheet.Copy(Before: newWorkbook.Worksheets[1]);
                    newWorkbook.Worksheets[2].Delete();

                    if (Path.GetExtension(sourceFile).ToLower() == ".xls")
                    {
                        newWorkbook.SaveAs(outputPath, Excel.XlFileFormat.xlExcel8);
                    }
                    else
                    {
                        newWorkbook.SaveAs(outputPath);
                    }

                    newWorkbook.Close(false);
                    ReleaseComObject(sheet);
                    ReleaseComObject(newWorkbook);

                    progressBar.Value = i;
                    lblProgress.Text = $"正在拆分: {i}/{sheetCount} - {sheetName}";
                    Application.DoEvents();
                }

                MessageBox.Show("拆分完成！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("发生错误：" + ex.Message);
            }
            finally
            {
                if (sourceWorkbook != null)
                {
                    sourceWorkbook.Close(false);
                    ReleaseComObject(sourceWorkbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    ReleaseComObject(excelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // 浏览按钮点击事件
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 文件 (*.xls;*.xlsx)|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = openFileDialog.FileName;
            }
        }

        // 释放 COM 对象的辅助方法
        private void ReleaseComObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch
            {
                obj = null;
            }
        }
    }
}
