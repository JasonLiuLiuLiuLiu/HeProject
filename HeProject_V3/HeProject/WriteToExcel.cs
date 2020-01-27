using System;
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

                ISheet sheet1 = workbook.CreateSheet("Sheet1");
                sheet1.DefaultColumnWidth = 1;
                sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
                SetHeader(sheet1, 0);
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
                    cell.SetCellValue(_context.GetP1Value<int>(7, rowIndex, i - beforeColumn));
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
                    cell.SetCellValue(_context.GetP1Value<int>(8, rowIndex, i - beforeColumn));
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
                    cell.SetCellValue(_context.GetP1Value<int>(9, rowIndex, i - beforeColumn));
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
                    cell.SetCellValue(_context.GetP1Value<int>(10, rowIndex, i - beforeColumn));
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
                    cell.SetCellValue(_context.GetP1Value<int>(11, rowIndex, i - beforeColumn));
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
                    cell.SetCellValue(_context.GetP1Value<int>(12, rowIndex, i - beforeColumn));
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