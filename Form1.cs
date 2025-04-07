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
            string sourceFile = txtFilePath.Text;  // 获取文件路径
            string outputFolder = Path.Combine(Path.GetDirectoryName(sourceFile), Path.GetFileNameWithoutExtension(sourceFile));  // 基于当前文件生成文件夹

            // 检查文件是否存在
            if (!File.Exists(sourceFile))
            {
                MessageBox.Show("源文件不存在！");
                return;
            }

            lblProgress.Text = $"正在拆分...";
            Application.DoEvents();  // 刷新界面，防止界面卡顿

            // 创建输出文件夹（如果不存在）
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // 创建 Excel 应用实例
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook sourceWorkbook = null;

            try
            {
                // 打开 Excel 文件
                if (Path.GetExtension(sourceFile).ToLower() == ".xls")
                {
                    sourceWorkbook = excelApp.Workbooks.Open(sourceFile, 0, false, 5);  // .xls文件
                }
                else if (Path.GetExtension(sourceFile).ToLower() == ".xlsx")
                {
                    sourceWorkbook = excelApp.Workbooks.Open(sourceFile, 0, false, 1);  // .xlsx文件
                }
                else
                {
                    MessageBox.Show("不支持的文件格式！");
                    return;
                }

                // 设置 Excel 的警告信息不弹出
                excelApp.DisplayAlerts = false;

                // 遍历工作表，拆分每个工作表
                int sheetCount = sourceWorkbook.Worksheets.Count;
                progressBar.Maximum = sheetCount; // 设置进度条最大值                

                for (int i = 1; i <= sheetCount; i++)
                {
                    Excel.Worksheet sheet = sourceWorkbook.Worksheets[i];
                    string sheetName = sheet.Name;
                    string outputPath = Path.Combine(outputFolder, $"{sheetName}{Path.GetExtension(sourceFile)}");

                    // 创建新工作簿
                    Excel.Workbook newWorkbook = excelApp.Workbooks.Add();

                    // 删除新工作簿中的默认 Sheet
                    while (newWorkbook.Worksheets.Count > 1)
                    {
                        newWorkbook.Worksheets[1].Delete();
                    }

                    // 拷贝当前工作表到新工作簿
                    sheet.Copy(Before: newWorkbook.Worksheets[1]);

                    // 删除新工作簿中多余的默认 Sheet
                    newWorkbook.Worksheets[2].Delete();

                    // 保存新文件
                    if (Path.GetExtension(sourceFile).ToLower() == ".xls")
                    {
                        newWorkbook.SaveAs(outputPath, Excel.XlFileFormat.xlExcel8);  // 保存为 .xls 格式
                    }
                    else
                    {
                        newWorkbook.SaveAs(outputPath);  // 保存为 .xlsx 格式
                    }

                    newWorkbook.Close(false);

                    // 更新进度条
                    progressBar.Value = i;
                    lblProgress.Text = $"正在拆分: {i}/{sheetCount} - {sheetName}";
                    Application.DoEvents();  // 刷新界面，防止界面卡顿
                }

                MessageBox.Show("拆分完成！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("发生错误：" + ex.Message);
            }
            finally
            {
                // 关闭源文件
                if (sourceWorkbook != null)
                {
                    sourceWorkbook.Close(false);
                }
                excelApp.Quit();
            }
        }

        // 浏览按钮点击事件，选择源文件
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            // 打开文件选择对话框
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 文件 (*.xls;*.xlsx)|*.xls;*.xlsx"; // 设置过滤器，只显示 Excel 文件
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // 将选中的文件路径显示到 TextBox 中
                txtFilePath.Text = openFileDialog.FileName;
            }
        }
    }
}
