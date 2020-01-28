using System;
using System.Collections.Generic;
using System.Diagnostics;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using HeProject.Model;
using NPOI.SS.Util;

namespace HeProject
{
    public class WriteToExcel
    {
        private readonly ProcessContext _context;
        private ICellStyle _headerStyle;
        private readonly Dictionary<int, ICellStyle> _styles = new Dictionary<int, ICellStyle>(60);
        private Dictionary<int, ProcessContext> _resultDic = new Dictionary<int, ProcessContext>(11);

        public WriteToExcel(Dictionary<int, ProcessContext> contexts)
        {
            _resultDic = contexts;
            _context = contexts[0];
        }

        public void Write()
        {
            try
            {
                IWorkbook workbook = new XSSFWorkbook();
                _headerStyle = workbook.CreateCellStyle();

                ISheet sheet1 = workbook.CreateSheet("总");
                sheet1.DefaultColumnWidth = 1;
                sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
                SetHeader(sheet1, 0);
                CreateStyles(workbook);
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

                    SetP2Value(row, index);
                }

                for (int i = 1; i < 11; i++)
                {
                    IRow row = sheet1.CreateRow(_context.Capacity + i * 2);

                    var context = _resultDic[i];

                    var index = context.Capacity - 1;

                    #region P1

                    SetP1S0Value(row, index, context);
                    SetP1S1Value(row, index, context);
                    SetP1S2Value(row, index, context);
                    SetP1S3Value(row, index, context);
                    SetP1S4Value(row, index, context);
                    SetP1S5Value(row, index, context);
                    SetP1S6Value(row, index, context);
                    SetP1S7Value(row, index, context);
                    SetP1S8Value(row, index, context);

                    #endregion P1

                    SetP2Value(row, index, context);
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

            var s0P1 = row.CreateCell(0);
            var s1P1 = row.CreateCell(StepLength.SourceLength * 1);
            var s2P1 = row.CreateCell(StepLength.SourceLength * 2);
            var s3P1 = row.CreateCell(StepLength.SourceLength * 3);
            var s4P1 = row.CreateCell(StepLength.SourceLength * 3 + 3);
            var s5P1 = row.CreateCell(StepLength.SourceLength * 3 + 6);
            var s6P1 = row.CreateCell(StepLength.SourceLength * 3 + 9);
            var s7P1 = row.CreateCell(StepLength.SourceLength * 3 + 12);
            var s8P1 = row.CreateCell(StepLength.SourceLength * 3 + 15);


            s0P1.SetCellValue("源数据");
            s1P1.SetCellValue("排序1");
            s2P1.SetCellValue("排序2");
            s3P1.SetCellValue("A");
            s4P1.SetCellValue("B");
            s5P1.SetCellValue("C");
            s6P1.SetCellValue("D");
            s7P1.SetCellValue("E");
            s8P1.SetCellValue("F");

            s0P1.CellStyle = _headerStyle;
            s1P1.CellStyle = _headerStyle;
            s2P1.CellStyle = _headerStyle;
            s3P1.CellStyle = _headerStyle;
            s4P1.CellStyle = _headerStyle;
            s5P1.CellStyle = _headerStyle;
            s6P1.CellStyle = _headerStyle;
            s7P1.CellStyle = _headerStyle;
            s8P1.CellStyle = _headerStyle;




            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 0, StepLength.SourceLength * 1 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 1, StepLength.SourceLength * 2 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 2, StepLength.SourceLength * 3 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3, StepLength.SourceLength * 3 + 2));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3 + 3, StepLength.SourceLength * 3 + 5));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3 + 6, StepLength.SourceLength * 3 + 8));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3 + 9, StepLength.SourceLength * 3 + 11));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3 + 12, StepLength.SourceLength * 3 + 14));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3 + 15, StepLength.SourceLength * 3 + 17));

            var columnWidth = 4;
            var beforeLength = StepLength.SourceLength * 3 + 18;

            var s0 = row.CreateCell(beforeLength + columnWidth * 0);
            var s1 = row.CreateCell(beforeLength + columnWidth * 1);
            var s2 = row.CreateCell(beforeLength + columnWidth * 2);
            var s3 = row.CreateCell(beforeLength + columnWidth * 3);
            var s4 = row.CreateCell(beforeLength + columnWidth * 4);
            var s5 = row.CreateCell(beforeLength + columnWidth * 5);
            var s6 = row.CreateCell(beforeLength + columnWidth * 6);
            var s7 = row.CreateCell(beforeLength + columnWidth * 7);
            var s8 = row.CreateCell(beforeLength + columnWidth * 8);
            var s9 = row.CreateCell(beforeLength + columnWidth * 9);
            var s10 = row.CreateCell(beforeLength + columnWidth * 10);
            var s11 = row.CreateCell(beforeLength + columnWidth * 11);
            var s12 = row.CreateCell(beforeLength + columnWidth * 12);
            var s13 = row.CreateCell(beforeLength + columnWidth * 13);
            var s14 = row.CreateCell(beforeLength + columnWidth * 14);
            var s15 = row.CreateCell(beforeLength + columnWidth * 15);
            var s16 = row.CreateCell(beforeLength + columnWidth * 16);
            var s17 = row.CreateCell(beforeLength + columnWidth * 17);
            var s18 = row.CreateCell(beforeLength + columnWidth * 18);
            var s19 = row.CreateCell(beforeLength + columnWidth * 19);
            var s20 = row.CreateCell(beforeLength + columnWidth * 20);
            var s21 = row.CreateCell(beforeLength + columnWidth * 21);
            var s22 = row.CreateCell(beforeLength + columnWidth * 22);
            var s23 = row.CreateCell(beforeLength + columnWidth * 23);
            var s24 = row.CreateCell(beforeLength + columnWidth * 24);
            var s25 = row.CreateCell(beforeLength + columnWidth * 25);
            var s26 = row.CreateCell(beforeLength + columnWidth * 26);
            var s27 = row.CreateCell(beforeLength + columnWidth * 27);
            var s28 = row.CreateCell(beforeLength + columnWidth * 28);
            var s29 = row.CreateCell(beforeLength + columnWidth * 29);
            var s30 = row.CreateCell(beforeLength + columnWidth * 30);
            var s31 = row.CreateCell(beforeLength + columnWidth * 31);
            var s32 = row.CreateCell(beforeLength + columnWidth * 32);

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
            s27.SetCellValue("单双1");
            s28.SetCellValue("单双2");
            s29.SetCellValue("大小1");
            s30.SetCellValue("大小2");
            s31.SetCellValue("阴阳1");
            s32.SetCellValue("阴阳2");

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
            s27.CellStyle = _headerStyle;
            s28.CellStyle = _headerStyle;
            s29.CellStyle = _headerStyle;
            s30.CellStyle = _headerStyle;
            s31.CellStyle = _headerStyle;
            s32.CellStyle = _headerStyle;

            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 0, beforeLength + columnWidth * 1 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 1, beforeLength + columnWidth * 2 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 2, beforeLength + columnWidth * 3 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 3, beforeLength + columnWidth * 4 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 4, beforeLength + columnWidth * 5 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 5, beforeLength + columnWidth * 6 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 6, beforeLength + columnWidth * 7 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 7, beforeLength + columnWidth * 8 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 8, beforeLength + columnWidth * 9 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 9, beforeLength + columnWidth * 10 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 10, beforeLength + columnWidth * 11 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 11, beforeLength + columnWidth * 12 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 12, beforeLength + columnWidth * 13 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 13, beforeLength + columnWidth * 14 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 14, beforeLength + columnWidth * 15 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 15, beforeLength + columnWidth * 16 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 16, beforeLength + columnWidth * 17 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 17, beforeLength + columnWidth * 18 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 18, beforeLength + columnWidth * 19 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 19, beforeLength + columnWidth * 20 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 20, beforeLength + columnWidth * 21 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 21, beforeLength + columnWidth * 22 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 22, beforeLength + columnWidth * 23 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 23, beforeLength + columnWidth * 24 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 24, beforeLength + columnWidth * 25 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 25, beforeLength + columnWidth * 26 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 26, beforeLength + columnWidth * 27 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 27, beforeLength + columnWidth * 28 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 28, beforeLength + columnWidth * 29 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 29, beforeLength + columnWidth * 30 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 30, beforeLength + columnWidth * 31 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 31, beforeLength + columnWidth * 32 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, beforeLength + columnWidth * 32, beforeLength + columnWidth * 33 - 1));
        }

        private void CreateStyles(IWorkbook workbook)
        {
            var colorStartIndex = 10;
            for (int i = 0; i < 60; i++)
            {
                var cellStyle = workbook.CreateCellStyle();
                cellStyle.FillForegroundColor = (short)(i + colorStartIndex++);
                cellStyle.FillPattern = FillPattern.SolidForeground;
                _styles.Add(i, cellStyle);
            }
        }

        private void SetP2Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            var columnWidth = 4;
            var beforeLength = StepLength.SourceLength * 3 + 18;

            for (int i = 0; i < 27; i++)
            {
                for (int j = 0; j < columnWidth; j++)
                {
                    var columnIndex = i * columnWidth + j + beforeLength;
                    var cell = row.CreateCell(columnIndex);
                    cell.CellStyle = _styles[i];
                    if (context.GetP2Value<bool>(i + 1, rowIndex, j))
                        cell.SetCellValue(j);
                }
            }

            for (int i = 27; i < 33; i++)
            {
                for (int j = 0; j < columnWidth; j++)
                {
                    var columnIndex = i * columnWidth + j + beforeLength;
                    var cell = row.CreateCell(columnIndex);
                    cell.CellStyle = _styles[i - 10];
                    cell.SetCellValue(context.GetP2Value<int>(i + 1, rowIndex, j));
                }
            }
        }

        #region SetP1Value

        private void SetP1S0Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = 0;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP1Value<bool>(0, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _styles[1];
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP1S1Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 1;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP1Value<bool>(1, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _styles[2];
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP1S2Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 2;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP1Value<bool>(2, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _styles[3];
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP1S3Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 3;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _styles[4];
                if (context.GetP1StepState(7, rowIndex))
                {
                    cell.SetCellValue(context.GetP1Value<int>(7, rowIndex, i - beforeColumn));
                }
            }
        }
        private void SetP1S4Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 3 + 3;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _styles[5];
                if (context.GetP1StepState(8, rowIndex))
                {
                    cell.SetCellValue(context.GetP1Value<int>(8, rowIndex, i - beforeColumn));
                }
            }
        }
        private void SetP1S5Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 3 + 6;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _styles[6];
                if (context.GetP1StepState(9, rowIndex))
                {
                    cell.SetCellValue(context.GetP1Value<int>(9, rowIndex, i - beforeColumn));
                }
            }
        }
        private void SetP1S6Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 3 + 9;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _styles[7];
                if (context.GetP1StepState(10, rowIndex))
                {
                    cell.SetCellValue(context.GetP1Value<int>(10, rowIndex, i - beforeColumn));
                }
            }
        }

        private void SetP1S7Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 3 + 12;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _styles[8];
                if (context.GetP1StepState(11, rowIndex))
                {
                    cell.SetCellValue(context.GetP1Value<int>(11, rowIndex, i - beforeColumn));
                }
            }
        }

        private void SetP1S8Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 3 + 15;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _styles[9];
                if (context.GetP1StepState(12, rowIndex))
                {
                    cell.SetCellValue(context.GetP1Value<int>(12, rowIndex, i - beforeColumn));
                }
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