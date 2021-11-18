using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace irRenamer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string folderPath = "";       

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Select_Btn_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            if (dlg.ShowDialog() == true)
            {
                string filename = dlg.FileName;
                folderPath = filename.Substring(0, filename.LastIndexOf('\\'));
                TextBox_Path.Text = folderPath;
                GetNames_Btn.IsEnabled = true;
                OpenExcel_Btn.IsEnabled = true;
            }
        }

        private void GetNames_Btn_Click(object sender, RoutedEventArgs e)
        {
            DirectoryInfo di = new DirectoryInfo(folderPath);
            List<ExcelData> list = new List<ExcelData>();
            // 获取所有文件
            foreach (var file in di.GetFiles())
            {
                ExcelData data = new ExcelData();
                data.type = 0;
                int dotIndex = file.Name.LastIndexOf('.');
                if (dotIndex == -1)
                {
                    data.oldName = file.Name;
                }
                else
                {
                    data.oldName = file.Name.Substring(0, dotIndex);
                    data.oldExt = file.Name.Substring(dotIndex + 1, file.Name.Length - dotIndex -1);
                }
                list.Add(data);
            }
            // 获取所有文件夹
            foreach (var folder in di.GetDirectories())
            {
                ExcelData data = new ExcelData();
                data.type = 1;
                data.oldName = folder.Name;
                data.oldExt = "文件夹";
                list.Add(data);
            }

            // 将数据写入Excel文件，并返回是否写入成功
            WriteExcel(folderPath, list);

        }
        private void OpenExcel_Btn_Click(object sender, RoutedEventArgs e)
        {
            string path = Directory.GetCurrentDirectory() + "\\" + "rename.xlsx";
            if (File.Exists(path))
            {
                Process.Start(path);
                Rename_Btn.IsEnabled = true;
            }
            else
            {
                MessageBox.Show("文件不存在，请重做上一步操作！");
            }

        }
        private void Rename_Btn_Click(object sender, RoutedEventArgs e)
        {
            GetNames_Btn.IsEnabled = false;
            OpenExcel_Btn.IsEnabled = false;
            Rename_Btn.IsEnabled = false;

            List<ExcelData> list = new List<ExcelData>();
            string path = Directory.GetCurrentDirectory() + "\\" + "rename.xlsx";
            list = ReadExcel(path);
            string resultStr = "操作记录：\n";
            int failedCount = 0;
            int totalCount = 0;
            foreach (var item in list)
            {
                string oldPath = "";
                string newPath = "";

                if (item.type == 0)
                {
                    if (string.IsNullOrEmpty(item.newName) && string.IsNullOrEmpty(item.newExt))
                    {
                        continue;
                    }
                    else
                    {
                        oldPath = item.oldName + "." + item.oldExt;
                        newPath = (string.IsNullOrEmpty(item.newName)) ? item.oldName : item.newName;
                        newPath += ".";
                        newPath += (string.IsNullOrEmpty(item.newExt)) ? item.oldExt : item.newExt;
                    }
                }
                else if (item.type == 1)
                {
                    if (string.IsNullOrEmpty(item.newName))
                    {
                        continue;
                    }
                    else
                    {
                        oldPath = item.oldName;
                        newPath = item.newName;
                    }
                }

                if (string.Equals(oldPath, newPath, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    Directory.Move(folderPath + "\\" + oldPath, folderPath + "\\" + newPath);
                    resultStr += "成功：将  " + oldPath + "  重命名为  " + newPath + "\n";
                }
                catch (Exception)
                {
                    failedCount++;
                    resultStr += "失败：将  " + oldPath + "  重命名为  " + newPath + "\n";
                    continue;
                }
                totalCount++;
            }
            resultStr += "\n\n阅读后请关闭本文件，避免出现写入错误！";
            WriteTxt(folderPath, resultStr);
            string showMsg = "操作完成，共" + totalCount.ToString() + "项，失败" + failedCount.ToString() + "项";
            MessageBox.Show(showMsg);
        }

        static private List<ExcelData> ReadExcel(string path)
        {
            List<ExcelData> list = new List<ExcelData>();
            try
            {
                using (var fs = new FileStream(path, FileMode.Open))
                {
                    IWorkbook wb = new XSSFWorkbook(fs);
                    ISheet sheet = wb.GetSheetAt(0);
                    //int columnCount = sheet.GetRow(0).LastCellNum;
                    //int rowCount = sheet.LastRowNum;

                    foreach (IRow row in sheet)
                    {
                        if (row.RowNum == 0)
                        {
                            continue;
                        }
                        ExcelData excelData = new ExcelData();
                        if (row.GetCell(0) != null)
                        {
                            if (string.IsNullOrEmpty(row.GetCell(0).ToString()))
                            {
                                continue;
                            }
                            else
                            {
                                excelData.oldName = row.GetCell(0).ToString();
                            }
                        }
                        if (row.GetCell(1) != null)
                        {
                            if (!string.IsNullOrEmpty(row.GetCell(1).ToString()))
                            {
                                excelData.oldExt = row.GetCell(1).ToString();
                                if (excelData.oldExt == "文件夹")
                                {
                                    excelData.type = 1;
                                }
                            }
                        }
                        if (row.GetCell(2) != null)
                        {
                            if (!string.IsNullOrEmpty(row.GetCell(2).ToString()))
                            {
                                excelData.newName = row.GetCell(2).ToString();
                            }
                        }
                        if (row.GetCell(3) != null)
                        {
                            if (!string.IsNullOrEmpty(row.GetCell(3).ToString()))
                            {
                                excelData.newExt = row.GetCell(3).ToString();
                            }
                        }
                        list.Add(excelData);
                    }
                    return list;
                }
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show("文件不存在");
                throw;
            }
            
        }
        static private bool WriteExcel(string folderPath, List<ExcelData> list)
        {
            string path = Directory.GetCurrentDirectory() + "\\" + "rename.xlsx";
            bool writeSuccess = false;

            try
            {
                using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    IWorkbook wb = new XSSFWorkbook();
                    ISheet sheet = wb.CreateSheet("Sheet1");

                    List<String> columns = new List<string>();
                    int rowIndex = 0;
                    // 写入表头
                    IRow row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue("原文件名");
                    row.CreateCell(1).SetCellValue("原扩展名");
                    row.CreateCell(2).SetCellValue("新文件名");
                    row.CreateCell(3).SetCellValue("新扩展名");
                    rowIndex++;

                    // 写入数据
                    foreach (var item in list)
                    {
                        row = sheet.CreateRow(rowIndex);
                        row.CreateCell(0).SetCellValue(item.oldName);
                        row.CreateCell(1).SetCellValue(item.oldExt);
                        //row.CreateCell(2).SetCellValue("newName");
                        //row.CreateCell(3).SetCellValue("newExt");
                        rowIndex++;
                    }
                    wb.Write(fs);
                }
                writeSuccess = true;
                MessageBox.Show("操作完成");
            }
            catch (Exception)
            {
                writeSuccess = false;
                MessageBox.Show("操作失败，请检查rename.xlsx文件是否已打开！");
            }
            return writeSuccess;
        }

        static private List<ExcelData> ReadCsv(string path)
        {
            ExcelData data = new ExcelData();
            List<ExcelData> list = new List<ExcelData>();
            using (StreamReader sr = new StreamReader(path))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    data.oldName = line.Split(',')[0];
                    data.oldExt = line.Split(',')[1];
                    data.newName = line.Split(',')[2];
                    data.newExt = line.Split(',')[3];
                    if (data.oldExt == "文件夹")
                    {
                        data.type = 1;
                    }
                    list.Add(data);
                }
            }
            return list;
        }

        static private void WriteCsv(string folderPath, List<ExcelData> list)
        {
            string path = Directory.GetCurrentDirectory() + "\\" + "rename.csv";

            FileStream fs = new FileStream(path, FileMode.OpenOrCreate);
            using (StreamWriter sw = new StreamWriter(fs, Encoding.UTF8))
            {
                string line = "原文件名,原扩展名,新文件名,新扩展名";
                sw.WriteLine(line);
                foreach (var item in list)
                {
                    line = item.oldName + "," + item.oldExt + "," + item.newName + "," + item.newExt;
                    sw.WriteLine(line);
                }
            }
            fs.Close();
        }

        static private string ReadTxt(string path)
        {
            string str = "";
            using (StreamReader sr = new StreamReader(path))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    str += line;
                }
            }
            return str;
        }

        static private void WriteTxt(string folderPath, string textstr)
        {
            string path = Directory.GetCurrentDirectory() + "\\" + "result.txt";

            FileStream fs = new FileStream(path, FileMode.Create);
            using (StreamWriter sw = new StreamWriter(fs, Encoding.UTF8))
            {
                sw.Write(textstr);
            }
            fs.Close();
        }

    }
}
