using System;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;
using System.Drawing;
using System.IO;

namespace Sofar.PointTableHelper.ExcelLib
{
    public class EPPlusHelpr
    {
        private EPPlusHelpr()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        private static void SetFont(Color _Color, ExcelRangeBase _Cells, string _FontName = "Arial", int _FontSize = 12)
        {
            ExcelRangeBase range = _Cells;
            ExcelStyle style = range.Style;
            // 修改字体名称  
            style.Font.Name = _FontName;
            // 修改字体大小  
            style.Font.Size = _FontSize;
            // 修改字体颜色（例如，设置为红色）  
            style.Font.Color.SetColor(_Color);
            // 修改字体加粗  
            style.Font.Bold = true;
        }
        private static void SetBack(Color _Color, ExcelRangeBase _Cells)
        {
            _Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            _Cells.Style.Fill.BackgroundColor.SetColor(_Color);
        }
        private static void SetBorder(ExcelRangeBase _Cells, ExcelBorderStyle _Style, Color _Color)
        {
            if (_Cells == null) return;

            _Cells.Style.Border.Top.Style = _Style;
            _Cells.Style.Border.Bottom.Style = _Style;
            _Cells.Style.Border.Left.Style = _Style;
            _Cells.Style.Border.Right.Style = _Style;

            _Cells.Style.Border.Top.Color.SetColor(_Color);
            _Cells.Style.Border.Bottom.Color.SetColor(_Color);
            _Cells.Style.Border.Left.Color.SetColor(_Color);
            _Cells.Style.Border.Right.Color.SetColor(_Color);

        }
        private static void SetBorder(ExcelWorksheet _WorkSheet, int _RowFrom, int _ColFrom, int _RowTo, int _ColTo, ExcelBorderStyle _Style, Color _Color)
        {
            ExcelRangeBase _Cells = GetExcelRangeBase(_WorkSheet, _RowFrom, _ColFrom, _RowTo, _ColTo);
            SetBorder(_Cells, _Style, _Color);
        }
        private static ExcelRangeBase GetExcelRangeBase(ExcelWorksheet _WorkSheet, int _RowFrom, int _ColFrom, int _RowTo, int _ColTo)
        {
            //_WorkSheet.Dimension.End.Column;
            try
            {
                ExcelAddressBase dimension = _WorkSheet.Dimension;
                ExcelRangeBase ret = _WorkSheet.Cells[_RowFrom, _ColFrom, _RowTo, _ColTo];
                return ret;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        private static ExcelRangeBase SetColData(ExcelWorksheet _WorkSheet, int _ColIdex, List<string> _DataList, int _RowFrom = 1)
        {
            try
            {
                int idex = _RowFrom;
                foreach (string _Header in _DataList)
                {
                    _WorkSheet.Cells[idex++, _ColIdex].Value = _Header;
                }

                ExcelRangeBase ret = GetExcelRangeBase(_WorkSheet, _RowFrom, _ColIdex, idex - 1, _ColIdex);
                return ret;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        private static ExcelRangeBase SetRowData(ExcelWorksheet _WorkSheet, int _RowIdex, List<string> _DataList, int _ColFrom = 1)
        {
            try
            {
                int idex = _ColFrom;
                foreach (string _Header in _DataList)
                {
                    _WorkSheet.Cells[_RowIdex, idex++].Value = _Header;
                }
                ExcelRangeBase ret = GetExcelRangeBase(_WorkSheet, _RowIdex, _ColFrom, _RowIdex, idex - 1);
                return ret;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        private static FileInfo GetFile(string _Name)
        {
            FileInfo excelFile = new FileInfo(_Name);
            return excelFile;
        }
        public static ExcelPackage GetPackage(string _FileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(GetFile(_FileName));
            return package;
        }
        public static ExcelPackage GetPackage()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage();
            return package;
        }
        private static ExcelPackage GetPackage(FileInfo _File)
        {
            ExcelPackage package = new ExcelPackage(_File);
            return package;
        }
        private static ExcelWorksheet GetWorksheet(ExcelPackage _Package, ref int _lastRowNumber, int _Sheet = 0)
        {
            ExcelWorksheet worksheet = _Package.Workbook.Worksheets[_Sheet];
            _lastRowNumber = worksheet.Dimension.End.Row;
            return worksheet;
        }
        private static ExcelWorksheet AddWorksheet(ExcelPackage _Package, string _WorkSheetName)
        {
            ExcelWorksheet worksheet = _Package.Workbook.Worksheets.Add(_WorkSheetName);
            return worksheet;
        }
        private static void SetWorkSheetStyle(ExcelWorksheet _WorkSheet)
        {
            _WorkSheet.Protection.IsProtected = false;
            _WorkSheet.Cells.AutoFitColumns();
        }





        public static void SetFileStyleA(string _FileName)
        {
            if (!File.Exists(_FileName))
            {
                return;
            }
            ExcelPackage package = EPPlusHelpr.GetPackage(_FileName);

            foreach (var sheet in package.Workbook.Worksheets)
            {
                ExcelWorksheet worksheet = sheet;
                ExcelRangeBase range1 = EPPlusHelpr.GetExcelRangeBase(worksheet, 1, 1, 1, worksheet.Dimension.End.Column);
                EPPlusHelpr.SetBack(Color.FromArgb(237, 237, 222), range1);
                EPPlusHelpr.SetFont(System.Drawing.Color.Blue, range1, "微软雅黑", 10);

                ExcelRangeBase range2 = EPPlusHelpr.GetExcelRangeBase(worksheet, 2, 1, worksheet.Dimension.End.Row, 1);
                EPPlusHelpr.SetBack(Color.FromArgb(200, 233, 237), range2);

                worksheet.Protection.IsProtected = false;
                worksheet.Protection.SetPassword(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                EPPlusHelpr.SetWorkSheetStyle(worksheet);

                EPPlusHelpr.SetBorder(worksheet, 1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns, ExcelBorderStyle.DashDot, Color.Black);
            }
            //ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            FileInfo excelFile = EPPlusHelpr.GetFile(DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xlsx");
            package.SaveAs(excelFile);
        }
        public static void SetFileStyleB(string _FileName)
        {
            if (!File.Exists(_FileName))
            {
                return;
            }
            ExcelPackage package = EPPlusHelpr.GetPackage(_FileName);

            foreach (var sheet in package.Workbook.Worksheets)
            {
                ExcelWorksheet worksheet = sheet;
                ExcelRangeBase range1 = EPPlusHelpr.GetExcelRangeBase(worksheet, 1, 1, 1, worksheet.Dimension.End.Column);
                EPPlusHelpr.SetBack(Color.FromArgb(237, 237, 222), range1);
                EPPlusHelpr.SetFont(System.Drawing.Color.Blue, range1, "微软雅黑", 10);

                ExcelRangeBase range2 = EPPlusHelpr.GetExcelRangeBase(worksheet, 2, 1, worksheet.Dimension.End.Row, 1);
                EPPlusHelpr.SetBack(Color.FromArgb(200, 233, 237), range2);

                worksheet.Protection.IsProtected = false;
                EPPlusHelpr.SetWorkSheetStyle(worksheet);

                EPPlusHelpr.SetBorder(worksheet, 1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns, ExcelBorderStyle.DashDot, Color.Black);
            }
            //ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            FileInfo excelFile = EPPlusHelpr.GetFile(DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xlsx");
            package.SaveAs(excelFile);
        }
        public static void CreatDemo()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // 创建一个新的 ExcelPackage 实例，这相当于一个 Excel 文件  
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = EPPlusHelpr.AddWorksheet(package, "测试页");

                List<string> list2 = new List<string>();
                list2.Add("Cow-1               1");
                list2.Add("Cow-2");
                list2.Add("Cow-3");
                list2.Add("Cow-4");
                list2.Add("Cow-5");
                ExcelRangeBase range1 = EPPlusHelpr.SetRowData(worksheet, 1, list2, 2);
                range1.Style.Locked = true;
                EPPlusHelpr.SetBack(System.Drawing.Color.Gray, range1);
                EPPlusHelpr.SetFont(System.Drawing.Color.Blue, range1, "微软雅黑", 10);


                List<string> list1 = new List<string>();
                list1.Add("Row-1");
                list1.Add("Row-2");
                list1.Add("Row-3");
                list1.Add("Row-4");
                list1.Add("Row-5");
                ExcelRangeBase range2 = EPPlusHelpr.SetColData(worksheet, 1, list1, 2);
                range2.Style.Locked = false;
                EPPlusHelpr.SetBack(System.Drawing.Color.Orange, range2);
                EPPlusHelpr.SetFont(System.Drawing.Color.Black, range2, "仿宋", 10);

                EPPlusHelpr.SetBorder(worksheet, 1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns, ExcelBorderStyle.DashDot, Color.Black);

                worksheet.Protection.IsProtected = false;
                // 如果设置了密码，则在这里提供密码字符串  
                worksheet.Protection.SetPassword(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                EPPlusHelpr.SetWorkSheetStyle(worksheet);
                FileInfo excelFile = EPPlusHelpr.GetFile("Demo.xlsx");
                package.SaveAs(excelFile);


            }

            // 输出文件保存成功的消息  
            Console.WriteLine("Excel file created successfully!");
        }
        public static void LoadFromDataTables(List<DataTable> _DataTableList, string _FileName = "")
        {
            ExcelPackage package = EPPlusHelpr.GetPackage();

            foreach (var dataTable in _DataTableList)
            {
                string temp = dataTable.TableName;
                if (temp == null || temp == "")
                    temp = DateTime.Now.ToString("dd-HH-mm-ss-fff");
                ExcelWorksheet worksheet = EPPlusHelpr.AddWorksheet(package, temp);

                worksheet.Cells["A1"].LoadFromDataTable(dataTable, PrintHeaders: true);


                ExcelRangeBase range1 = EPPlusHelpr.GetExcelRangeBase(worksheet, 1, 1, 1, worksheet.Dimension.End.Column);
                EPPlusHelpr.SetBack(Color.FromArgb(237, 237, 222), range1);
                EPPlusHelpr.SetFont(System.Drawing.Color.Blue, range1, "微软雅黑", 10);

                ExcelRangeBase range2 = EPPlusHelpr.GetExcelRangeBase(worksheet, 2, 1, worksheet.Dimension.End.Row, 1);
                EPPlusHelpr.SetBack(Color.FromArgb(200, 233, 237), range2);

                worksheet.Protection.IsProtected = false;
                worksheet.Protection.SetPassword(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                EPPlusHelpr.SetWorkSheetStyle(worksheet);

                EPPlusHelpr.SetBorder(worksheet, 1, 1, worksheet.Dimension.Rows, worksheet.Dimension.Columns, ExcelBorderStyle.DashDot, Color.Black);
            }
            //ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            FileInfo excelFile = EPPlusHelpr.GetFile(_FileName == "" ? (DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xlsx") : _FileName);

            package.SaveAs(excelFile);
        }
        public static List<DataTable> LoadFromFile(string _FileName, bool _HasHeader = true)
        {
            if (!File.Exists(_FileName))
            {
                return new List<DataTable>();
            }
            ExcelPackage package = EPPlusHelpr.GetPackage();
            try
            {
                package = EPPlusHelpr.GetPackage(_FileName);
                List<DataTable> DataTableList = new List<DataTable>();

                foreach (var worksheet in package.Workbook.Worksheets)
                {

                    DataTable dataTable = new DataTable();

                    int rowStart = 1;

                    if (_HasHeader == true)
                    {
                        foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                        {
                            dataTable.Columns.Add(firstRowCell.Text);
                        }

                        // 遍历剩余的行以填充DataTable的行  
                        rowStart = 2; // 假设第一行是标题行，从第二行开始读取数据  
                    }

                    //while (worksheet.Cells[rowStart, 1].Value != null)
                    //{
                    //    DataRow newRow = dataTable.NewRow();
                    //    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    //    {
                    //        ExcelRangeBase cell = worksheet.Cells[rowStart, col];
                    //        if (cell.Value != null)
                    //        {
                    //            newRow[col - 1] = cell.Value; // 列索引从0开始，而Excel的列从1开始，所以减1  
                    //        }
                    //    }
                    //    dataTable.Rows.Add(newRow);
                    //    rowStart++;
                    //}
                    while (worksheet.Cells[rowStart, 1].Value != null)
                    {
                        DataRow newRow = dataTable.NewRow();
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            ExcelRangeBase cell = worksheet.Cells[rowStart, col];
                            if (cell.Value != null)
                            {
                                if (cell.Value is DateTime)
                                {
                                    newRow[col - 1] = ((DateTime)cell.Value).ToString("yyyy-MM-dd");
                                }
                                else
                                {
                                    newRow[col - 1] = cell.Text;
                                }
                            }
                        }
                        dataTable.Rows.Add(newRow);
                        rowStart++;
                    }

                    DataTableList.Add(dataTable);
                }
                package.Dispose();
                return DataTableList;

            }
            catch (Exception ex)
            {
                package.Dispose();
                return new List<DataTable>();
            }
        }
    }
}
