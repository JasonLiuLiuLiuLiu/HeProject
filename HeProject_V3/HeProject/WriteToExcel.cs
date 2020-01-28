using System;
using System.Collections.Generic;
using System.Diagnostics;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Linq;
using HeProject.Model;
using NPOI.SS.Util;

namespace HeProject
{
    public class WriteToExcel
    {
        private readonly ProcessContext _context;
        private ICellStyle _headerStyle;
        private Dictionary<int, ISheet> sheets = new Dictionary<int, ISheet>(6);
        private Dictionary<int, ICellStyle> styles = new Dictionary<int, ICellStyle>(50);

        private ICellStyle _s0P1Style;
        private ICellStyle _s1P1Style;
        private ICellStyle _s2P1Style;
        private ICellStyle _s3P1Style;
        private ICellStyle _s4P1Style;
        private ICellStyle _s5P1Style;
        private ICellStyle _s6P1Style;
        private ICellStyle _s7P1Style;
        private ICellStyle _s8P1Style;


        public WriteToExcel(ProcessContext context)
        {
            _context = context;
        }

        public void Write()
        {
            try
            {
                IWorkbook workbook = new XSSFWorkbook();
                _headerStyle = workbook.CreateCellStyle();

                #region SetStyle

                _s0P1Style = workbook.CreateCellStyle();
                _s1P1Style = workbook.CreateCellStyle();
                _s2P1Style = workbook.CreateCellStyle();
                _s3P1Style = workbook.CreateCellStyle();
                _s4P1Style = workbook.CreateCellStyle();
                _s5P1Style = workbook.CreateCellStyle();
                _s6P1Style = workbook.CreateCellStyle();
                _s7P1Style = workbook.CreateCellStyle();
                _s8P1Style = workbook.CreateCellStyle();


                //https://www.cnblogs.com/love-zf/p/4874098.html
                _s0P1Style.FillForegroundColor = 13;
                _s1P1Style.FillForegroundColor = 14;
                _s2P1Style.FillForegroundColor = 15;
                _s3P1Style.FillForegroundColor = 16;
                _s4P1Style.FillForegroundColor = 17;
                _s5P1Style.FillForegroundColor = 18;
                _s6P1Style.FillForegroundColor = 19;
                _s7P1Style.FillForegroundColor = 20;
                _s8P1Style.FillForegroundColor = 21;


                _s0P1Style.FillPattern = FillPattern.SolidForeground;
                _s1P1Style.FillPattern = FillPattern.SolidForeground;
                _s2P1Style.FillPattern = FillPattern.SolidForeground;
                _s3P1Style.FillPattern = FillPattern.SolidForeground;
                _s4P1Style.FillPattern = FillPattern.SolidForeground;
                _s5P1Style.FillPattern = FillPattern.SolidForeground;
                _s6P1Style.FillPattern = FillPattern.SolidForeground;
                _s7P1Style.FillPattern = FillPattern.SolidForeground;
                _s8P1Style.FillPattern = FillPattern.SolidForeground;



                #endregion SetStyle

                ISheet sheet1 = workbook.CreateSheet("总");
                sheet1.DefaultColumnWidth = 1;
                sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
                SetHeader(sheet1, 0);
                CreateSheetsForP2(workbook);
                CreateSheetsHeaderForP2();
                CreateStyleForP2(workbook);
                for (int i = 1; i <= _context.Capacity; i++)
                {
                    IRow row = sheet1.CreateRow(i);
                    var index = i - 1;

                    #region P1

                    SetP1S0Value(row, index);
                    SetP1S1Value(row, index);
                    SetP1S2Value(row, index);
                    SetP1S3Value(row, index);
                    SetP1S4Value(row, index);
                    SetP1S5Value(row, index);
                    SetP1S6Value(row, index);
                    SetP1S7Value(row, index);
                    SetP1S8Value(row, index);

                    #endregion P1
                }
                SaveFile(workbook);
            }
            catch (Exception ex)
            {
                Console.WriteLine("处理失败!");
                Console.WriteLine(ex.Message);
            }
        }

        private void SetHeader(ISheet sheet, int headerIndex)
        {
            var row = sheet.GetRow(headerIndex);
            _headerStyle.Alignment = HorizontalAlignment.Center;

            var s0p1 = row.CreateCell(0);
            var s1p1 = row.CreateCell(StepLength.SourceLength * 1);
            var s2p1 = row.CreateCell(StepLength.SourceLength * 2);
            var s3p1 = row.CreateCell(StepLength.SourceLength * 3);
            var s4p1 = row.CreateCell(StepLength.SourceLength * 3 + 3);
            var s5p1 = row.CreateCell(StepLength.SourceLength * 3 + 6);
            var s6p1 = row.CreateCell(StepLength.SourceLength * 3 + 9);
            var s7p1 = row.CreateCell(StepLength.SourceLength * 3 + 12);
            var s8p1 = row.CreateCell(StepLength.SourceLength * 3 + 15);

            s0p1.SetCellValue("源数据");
            s1p1.SetCellValue("排序1");
            s2p1.SetCellValue("排序2");
            s3p1.SetCellValue("A");
            s4p1.SetCellValue("B");
            s5p1.SetCellValue("C");
            s6p1.SetCellValue("D");
            s7p1.SetCellValue("E");
            s8p1.SetCellValue("F");

            s0p1.CellStyle = _headerStyle;
            s1p1.CellStyle = _headerStyle;
            s2p1.CellStyle = _headerStyle;
            s3p1.CellStyle = _headerStyle;
            s4p1.CellStyle = _headerStyle;
            s5p1.CellStyle = _headerStyle;
            s6p1.CellStyle = _headerStyle;
            s7p1.CellStyle = _headerStyle;
            s8p1.CellStyle = _headerStyle;


            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 0, StepLength.SourceLength * 1 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 1, StepLength.SourceLength * 2 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 2, StepLength.SourceLength * 3 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3, StepLength.SourceLength * 3 + 2));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3 + 3, StepLength.SourceLength * 3 + 5));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3 + 6, StepLength.SourceLength * 3 + 8));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3 + 9, StepLength.SourceLength * 3 + 11));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3 + 12, StepLength.SourceLength * 3 + 14));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3 + 15, StepLength.SourceLength * 3 + 17));
        }

        private void CreateSheetsForP2(IWorkbook workbook)
        {
            var a = workbook.CreateSheet("A");
            var b = workbook.CreateSheet("B");
            var c = workbook.CreateSheet("C");
            var d = workbook.CreateSheet("D");
            var e = workbook.CreateSheet("E");
            var f = workbook.CreateSheet("F");
            var g = workbook.CreateSheet("G");
            a.DefaultColumnWidth = 1;
            b.DefaultColumnWidth = 1;
            c.DefaultColumnWidth = 1;
            d.DefaultColumnWidth = 1;
            e.DefaultColumnWidth = 1;
            f.DefaultColumnWidth = 1;
            g.DefaultColumnWidth = 1;
            sheets.Add(0, a);
            sheets.Add(1, b);
            sheets.Add(2, c);
            sheets.Add(3, d);
            sheets.Add(4, e);
            sheets.Add(5, f);
            sheets.Add(6, g);
        }

        private void CreateSheetsHeaderForP2()
        {
            var headerIndex = 0;
            var columnWidth = 4;
            foreach (var sheet in sheets.Values)
            {
                var row = sheet.CreateRow(headerIndex);
                _headerStyle.Alignment = HorizontalAlignment.Center;
                var source = row.CreateCell(0);
                var s0 = row.CreateCell(3 + columnWidth * 0);
                var s1 = row.CreateCell(3 + columnWidth * 1);
                var s2 = row.CreateCell(3 + columnWidth * 2);
                var s3 = row.CreateCell(3 + columnWidth * 3);
                var s4 = row.CreateCell(3 + columnWidth * 4);
                var s5 = row.CreateCell(3 + columnWidth * 5);
                var s6 = row.CreateCell(3 + columnWidth * 6);
                var s7 = row.CreateCell(3 + columnWidth * 7);
                var s8 = row.CreateCell(3 + columnWidth * 8);
                var s9 = row.CreateCell(3 + columnWidth * 9);
                var s10 = row.CreateCell(3 + columnWidth * 10);
                var s11 = row.CreateCell(3 + columnWidth * 11);
                var s12 = row.CreateCell(3 + columnWidth * 12);
                var s13 = row.CreateCell(3 + columnWidth * 13);
                var s14 = row.CreateCell(3 + columnWidth * 14);
                var s15 = row.CreateCell(3 + columnWidth * 15);
                var s16 = row.CreateCell(3 + columnWidth * 16);
                var s17 = row.CreateCell(3 + columnWidth * 17);
                var s18 = row.CreateCell(3 + columnWidth * 18);
                var s19 = row.CreateCell(3 + columnWidth * 19);
                var s20 = row.CreateCell(3 + columnWidth * 20);
                var s21 = row.CreateCell(3 + columnWidth * 21);
                var s22 = row.CreateCell(3 + columnWidth * 22);
                var s23 = row.CreateCell(3 + columnWidth * 23);
                var s24 = row.CreateCell(3 + columnWidth * 24);
                var s25 = row.CreateCell(3 + columnWidth * 25);
                var s26 = row.CreateCell(3 + columnWidth * 26);

                source.SetCellValue("源数据");
                s0.SetCellValue("单双12");
                s1.SetCellValue("排序1");
                s2.SetCellValue("排序2");
                s3.SetCellValue("单双13");
                s4.SetCellValue("排序1");
                s5.SetCellValue("排序2");
                s6.SetCellValue("单双23");
                s7.SetCellValue("排序1");
                s8.SetCellValue("排序2");
                s9.SetCellValue("大小12");
                s10.SetCellValue("排序1");
                s11.SetCellValue("排序2");
                s12.SetCellValue("大小13");
                s13.SetCellValue("排序1");
                s14.SetCellValue("排序2");
                s15.SetCellValue("大小23");
                s16.SetCellValue("排序1");
                s17.SetCellValue("排序2");
                s18.SetCellValue("阴阳12");
                s19.SetCellValue("排序1");
                s20.SetCellValue("排序2");
                s21.SetCellValue("阴阳13");
                s22.SetCellValue("排序1");
                s23.SetCellValue("排序2");
                s24.SetCellValue("阴阳23");
                s25.SetCellValue("排序1");
                s26.SetCellValue("排序2");

                source.CellStyle = _headerStyle;
                s0.CellStyle = _headerStyle;
                s1.CellStyle = _headerStyle;
                s2.CellStyle = _headerStyle;
                s3.CellStyle = _headerStyle;
                s4.CellStyle = _headerStyle;
                s5.CellStyle = _headerStyle;
                s6.CellStyle = _headerStyle;
                s7.CellStyle = _headerStyle;
                s8.CellStyle = _headerStyle;
                s9.CellStyle = _headerStyle;
                s10.CellStyle = _headerStyle;
                s11.CellStyle = _headerStyle;
                s12.CellStyle = _headerStyle;
                s13.CellStyle = _headerStyle;
                s14.CellStyle = _headerStyle;
                s15.CellStyle = _headerStyle;
                s16.CellStyle = _headerStyle;
                s17.CellStyle = _headerStyle;
                s18.CellStyle = _headerStyle;
                s19.CellStyle = _headerStyle;
                s20.CellStyle = _headerStyle;
                s21.CellStyle = _headerStyle;
                s22.CellStyle = _headerStyle;
                s23.CellStyle = _headerStyle;
                s24.CellStyle = _headerStyle;
                s25.CellStyle = _headerStyle;
                s26.CellStyle = _headerStyle;

                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 0, 2));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 0, 3 + columnWidth * 1 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 1, 3 + columnWidth * 2 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 2, 3 + columnWidth * 3 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 3, 3 + columnWidth * 4 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 4, 3 + columnWidth * 5 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 5, 3 + columnWidth * 6 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 6, 3 + columnWidth * 7 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 7, 3 + columnWidth * 8 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 8, 3 + columnWidth * 9 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 9, 3 + columnWidth * 10 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 10, 3 + columnWidth * 11 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 11, 3 + columnWidth * 12 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 12, 3 + columnWidth * 13 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 13, 3 + columnWidth * 14 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 14, 3 + columnWidth * 15 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 15, 3 + columnWidth * 16 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 16, 3 + columnWidth * 17 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 17, 3 + columnWidth * 18 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 18, 3 + columnWidth * 19 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 19, 3 + columnWidth * 20 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 20, 3 + columnWidth * 21 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 21, 3 + columnWidth * 22 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 22, 3 + columnWidth * 23 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 23, 3 + columnWidth * 24 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 24, 3 + columnWidth * 25 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 25, 3 + columnWidth * 26 - 1));
                sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 3 + columnWidth * 26, 3 + columnWidth * 27 - 1));
            }
        }

        private void CreateStyleForP2(IWorkbook workbook)
        {
            var colorStartIndex = 10;
            for (int i = 0; i < 28; i++)
            {
                var cellStyle = workbook.CreateCellStyle();
                cellStyle.FillForegroundColor = (short)(i + colorStartIndex++);
                cellStyle.FillPattern = FillPattern.SolidForeground;
                styles.Add(i, cellStyle);
            }
        }

        private IRow GetSheetRow(int sheetIndex, int rowIndex)
        {
            var row = sheets[sheetIndex].GetRow(rowIndex);

            if (row == null)
            {
                row = sheets[sheetIndex].CreateRow(rowIndex);
            }

            return row;
        }

        private void SetRowValueForP2(IRow row, int rowIndex, int stage)
        {
            int columnWidth = 4;
            for (int i = 0; i < 3; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = styles[5];
                if (_context.GetP1StepState(7 + stage, rowIndex))
                {
                    cell.SetCellValue(_context.GetP1Value<int>(7 + stage, rowIndex, i));
                }
            }
            for (int i = 0; i < 27; i++)
            {
                for (int j = 0; j < columnWidth; j++)
                {
                    var columnIndex = i * columnWidth + j + 3;
                    var cell = row.CreateCell(columnIndex);
                    cell.CellStyle = styles[i];
                    if (_context.GetP2StepState(stage, i + 1, rowIndex) && _context.GetP2Value<bool>(stage, i + 1, rowIndex, j))
                        cell.SetCellValue(j);
                }
            }
        }

        private void SetRowValueForP2(int stage, int rowIndex)
        {
            SetRowValueForP2(GetSheetRow(stage, rowIndex + 1), rowIndex, stage);
        }

        #region SetP1Value

        private void SetP1S0Value(IRow row, int rowIndex)
        {
            int beforeColumn = 0;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP1Value<bool>(0, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s0P1Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP1S1Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 1;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP1Value<bool>(1, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s1P1Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP1S2Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 2;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP1Value<bool>(2, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s2P1Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP1S3Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 3;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s3P1Style;
                if (_context.GetP1StepState(7, rowIndex))
                {
                    cell.SetCellValue(_context.GetP1Value<int>(7, rowIndex, i - beforeColumn));
                }
            }
            if (_context.GetP1StepState(7, rowIndex))
            {
                SetRowValueForP2(0, rowIndex);
            }
        }
        private void SetP1S4Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 3 + 3;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s4P1Style;
                if (_context.GetP1StepState(8, rowIndex))
                {
                    cell.SetCellValue(_context.GetP1Value<int>(8, rowIndex, i - beforeColumn));
                }
            }
            if (_context.GetP1StepState(8, rowIndex))
            {
                SetRowValueForP2(1, rowIndex);
            }
        }
        private void SetP1S5Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 3 + 6;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s5P1Style;
                if (_context.GetP1StepState(9, rowIndex))
                {
                    cell.SetCellValue(_context.GetP1Value<int>(9, rowIndex, i - beforeColumn));
                }
            }
            if (_context.GetP1StepState(9, rowIndex))
            {
                SetRowValueForP2(2, rowIndex);
            }
        }
        private void SetP1S6Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 3 + 9;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s6P1Style;
                if (_context.GetP1StepState(10, rowIndex))
                {
                    cell.SetCellValue(_context.GetP1Value<int>(10, rowIndex, i - beforeColumn));
                }
            }
            if (_context.GetP1StepState(10, rowIndex))
            {
                SetRowValueForP2(3, rowIndex);
            }
        }

        private void SetP1S7Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 3 + 12;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s7P1Style;
                if (_context.GetP1StepState(11, rowIndex))
                {
                    cell.SetCellValue(_context.GetP1Value<int>(11, rowIndex, i - beforeColumn));
                }
            }
            if (_context.GetP1StepState(11, rowIndex))
            {
                SetRowValueForP2(4, rowIndex);
            }
        }

        private void SetP1S8Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 3 + 15;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s8P1Style;
                if (_context.GetP1StepState(12, rowIndex))
                {
                    cell.SetCellValue(_context.GetP1Value<int>(12, rowIndex, i - beforeColumn));
                }
            }
            if (_context.GetP1StepState(12, rowIndex))
            {
                SetRowValueForP2(5, rowIndex);
            }
        }
        #endregion SetP1Value


        private void SaveFile(IWorkbook workBook)
        {
            try
            {
                var path = AppDomain.CurrentDomain.BaseDirectory + DateTime.Now.ToString("yyyy-MM-dd");
                var fullPath = path + "\\" + DateTime.Now.ToString("HH-mm-ss-ffff") + ".xlsx";
                if (!Directory.Exists(path))//如果不存在就创建file文件夹　　             　　
                    Directory.CreateDirectory(path);//创建该文件夹　　
                FileStream sw = File.Create(fullPath);
                workBook.Write(sw);
                sw.Close();
                // Process.Start("explorer.exe", "/select, \"" + fullPath + "\"");
                Process.Start(fullPath);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.ReadKey();
                throw;
            }
        }
    }
}