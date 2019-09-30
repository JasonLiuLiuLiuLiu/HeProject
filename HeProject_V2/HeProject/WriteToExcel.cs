using System;
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

        private ICellStyle _s1P1Style;
        private ICellStyle _s2P1Style;
        private ICellStyle _s3P1Style;

        private ICellStyle _s1P2Style;
        private ICellStyle _s2P2Style;
        private ICellStyle _s3P2Style;

        private ICellStyle _s1P3Style;
        private ICellStyle _s2P3Style;
        private ICellStyle _s3P3Style;

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

                _s1P1Style = workbook.CreateCellStyle();
                _s2P1Style = workbook.CreateCellStyle();
                _s3P1Style = workbook.CreateCellStyle();

                _s1P2Style = workbook.CreateCellStyle();
                _s2P2Style = workbook.CreateCellStyle();
                _s3P2Style = workbook.CreateCellStyle();

                _s1P3Style = workbook.CreateCellStyle();
                _s2P3Style = workbook.CreateCellStyle();
                _s3P3Style = workbook.CreateCellStyle();

                //https://www.cnblogs.com/love-zf/p/4874098.html
                _s1P1Style.FillForegroundColor = 13;
                _s2P1Style.FillForegroundColor = 22;
                _s3P1Style.FillForegroundColor = 45;

                _s1P2Style.FillForegroundColor = 13;
                _s2P2Style.FillForegroundColor = 22;
                _s3P2Style.FillForegroundColor = 45;

                _s1P3Style.FillForegroundColor = 13;
                _s2P3Style.FillForegroundColor = 22;
                _s3P3Style.FillForegroundColor = 45;

                _s1P1Style.FillPattern = FillPattern.SolidForeground;
                _s2P1Style.FillPattern = FillPattern.SolidForeground;
                _s3P1Style.FillPattern = FillPattern.SolidForeground;

                _s1P2Style.FillPattern = FillPattern.SolidForeground;
                _s2P2Style.FillPattern = FillPattern.SolidForeground;
                _s3P2Style.FillPattern = FillPattern.SolidForeground;

                _s1P3Style.FillPattern = FillPattern.SolidForeground;
                _s2P3Style.FillPattern = FillPattern.SolidForeground;
                _s3P3Style.FillPattern = FillPattern.SolidForeground;

                #endregion SetStyle

                ISheet sheet1 = workbook.CreateSheet("Sheet1");
                sheet1.DefaultColumnWidth = 1;
                sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
                SetHeader(sheet1, 0);
                for (int i = 1; i <= _context.Capacity; i++)
                {
                    IRow row = sheet1.CreateRow(i);
                    var index = i-1;

                    #region P1

                    SetP1S1Value(row, index);
                    SetP1S2Value(row, index);
                    SetP1S3Value(row, index);

                    #endregion P1

                    #region P2

                    SetP2S1Value(row, index);
                    SetP2S2Value(row, index);
                    SetP2S3Value(row, index);

                    #endregion P2

                    #region P3

                    SetP3S1Value(row, index);
                    SetP3S2Value(row, index);
                    SetP3S3Value(row, index);

                    #endregion P3

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

            var s1p1 = row.CreateCell(0);
            var s2p1 = row.CreateCell(StepLength.SourceLength * 1);
            var s3p1 = row.CreateCell(StepLength.SourceLength * 2);

            s1p1.SetCellValue("A1");
            s2p1.SetCellValue("A2");
            s3p1.SetCellValue("A3");

            s1p1.CellStyle = _headerStyle;
            s2p1.CellStyle = _headerStyle;
            s3p1.CellStyle = _headerStyle;

            var s1p2 = row.CreateCell(StepLength.SourceLength * 3);
            var s2p2 = row.CreateCell(StepLength.SourceLength * 4);
            var s3p2 = row.CreateCell(StepLength.SourceLength * 5);

            s1p2.SetCellValue("B1");
            s2p2.SetCellValue("B2");
            s3p2.SetCellValue("B3");

            s1p2.CellStyle = _headerStyle;
            s2p2.CellStyle = _headerStyle;
            s3p2.CellStyle = _headerStyle;

            var s1p3 = row.CreateCell(StepLength.SourceLength * 6);
            var s2p3 = row.CreateCell(StepLength.SourceLength * 7);
            var s3p3 = row.CreateCell(StepLength.SourceLength * 8);

            s1p3.SetCellValue("C1");
            s2p3.SetCellValue("C2");
            s3p3.SetCellValue("C3");

            s1p3.CellStyle = _headerStyle;
            s2p3.CellStyle = _headerStyle;
            s3p3.CellStyle = _headerStyle;

            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 0, StepLength.SourceLength*1 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 1, StepLength.SourceLength * 2 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 2, StepLength.SourceLength * 3 - 1));

            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3 + 0, StepLength.SourceLength * 4 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 4, StepLength.SourceLength * 5 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 5, StepLength.SourceLength * 6 - 1));

            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 6 + 0, StepLength.SourceLength * 7 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 7, StepLength.SourceLength * 8 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 8, StepLength.SourceLength * 9 - 1));

        }

        #region SetP1Value



        private void SetP1S1Value(IRow row, int rowIndex)
        {
            int beforeColumn = 0;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP1Value<bool>(1, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s1P1Style;
                if (value)
                    cell.SetCellValue(i% StepLength.SourceLength);
            }
        }



        private void SetP1S2Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 1;
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
            int beforeColumn = StepLength.SourceLength * 2;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP1Value<bool>(3, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s3P1Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        #endregion SetP1Value

        #region SetP2Value

        private void SetP2S1Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 3;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP2Value<bool>(1, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s1P2Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP2S2Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 4;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP2Value<bool>(2, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s2P2Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP2S3Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 5;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP2Value<bool>(3, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s3P2Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        #endregion SetP2Value

        #region SetP3Value

        private void SetP3S1Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 6;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP3Value<bool>(1, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s1P3Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP3S2Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 7;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP3Value<bool>(2, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s2P3Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP3S3Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 8;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP3Value<bool>(3, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s3P3Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        #endregion SetP3Value


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