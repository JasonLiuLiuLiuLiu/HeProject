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

        private ICellStyle _s1P4Style;
        private ICellStyle _s2P4Style;
        private ICellStyle _s3P4Style;
        private ICellStyle _s4P4Style;
        private ICellStyle _s5P4Style;
        private ICellStyle _s6P4Style;
        private ICellStyle _s7P4Style;
        private ICellStyle _s8P4Style;
        private ICellStyle _s9P4Style;
        private ICellStyle _s10P4Style;
        private ICellStyle _s11P4Style;
        private ICellStyle _s12P4Style;
        private ICellStyle _s13P4Style;
        private ICellStyle _s14P4Style;
        private ICellStyle _s15P4Style;
        private ICellStyle _s16P4Style;
        private ICellStyle _s17P4Style;
        private ICellStyle _s18P4Style;
        private ICellStyle _s19P4Style;
        private ICellStyle _s20P4Style;
        private ICellStyle _s21P4Style;
        private ICellStyle _s22P4Style;
        private ICellStyle _s23P4Style;
        private ICellStyle _s24P4Style;

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

                _s1P4Style = workbook.CreateCellStyle();
                _s2P4Style = workbook.CreateCellStyle();
                _s3P4Style = workbook.CreateCellStyle();
                _s4P4Style = workbook.CreateCellStyle();
                _s5P4Style = workbook.CreateCellStyle();
                _s6P4Style = workbook.CreateCellStyle();
                _s7P4Style = workbook.CreateCellStyle();
                _s8P4Style = workbook.CreateCellStyle();
                _s9P4Style = workbook.CreateCellStyle();
                _s10P4Style = workbook.CreateCellStyle();
                _s11P4Style = workbook.CreateCellStyle();
                _s12P4Style = workbook.CreateCellStyle();
                _s13P4Style = workbook.CreateCellStyle();
                _s14P4Style = workbook.CreateCellStyle();
                _s15P4Style = workbook.CreateCellStyle();
                _s16P4Style = workbook.CreateCellStyle();
                _s17P4Style = workbook.CreateCellStyle();
                _s18P4Style = workbook.CreateCellStyle();
                _s19P4Style = workbook.CreateCellStyle();
                _s20P4Style = workbook.CreateCellStyle();
                _s21P4Style = workbook.CreateCellStyle();
                _s22P4Style = workbook.CreateCellStyle();
                _s23P4Style = workbook.CreateCellStyle();
                _s24P4Style = workbook.CreateCellStyle();

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

                _s1P4Style.FillForegroundColor = 13;
                _s2P4Style.FillForegroundColor = 22;
                _s3P4Style.FillForegroundColor = 45;

                _s4P4Style.FillForegroundColor = 13;
                _s5P4Style.FillForegroundColor = 22;
                _s6P4Style.FillForegroundColor = 45;

                _s7P4Style.FillForegroundColor = 13;
                _s8P4Style.FillForegroundColor = 22;
                _s9P4Style.FillForegroundColor = 45;

                _s10P4Style.FillForegroundColor = 13;
                _s11P4Style.FillForegroundColor = 22;
                _s12P4Style.FillForegroundColor = 45;

                _s13P4Style.FillForegroundColor = 13;
                _s14P4Style.FillForegroundColor = 22;
                _s15P4Style.FillForegroundColor = 45;

                _s16P4Style.FillForegroundColor = 13;
                _s17P4Style.FillForegroundColor = 22;
                _s18P4Style.FillForegroundColor = 45;

                _s19P4Style.FillForegroundColor = 13;
                _s20P4Style.FillForegroundColor = 22;
                _s21P4Style.FillForegroundColor = 45;

                _s22P4Style.FillForegroundColor = 13;
                _s23P4Style.FillForegroundColor = 22;
                _s24P4Style.FillForegroundColor = 45;

                _s1P1Style.FillPattern = FillPattern.SolidForeground;
                _s2P1Style.FillPattern = FillPattern.SolidForeground;
                _s3P1Style.FillPattern = FillPattern.SolidForeground;

                _s1P2Style.FillPattern = FillPattern.SolidForeground;
                _s2P2Style.FillPattern = FillPattern.SolidForeground;
                _s3P2Style.FillPattern = FillPattern.SolidForeground;

                _s1P3Style.FillPattern = FillPattern.SolidForeground;
                _s2P3Style.FillPattern = FillPattern.SolidForeground;
                _s3P3Style.FillPattern = FillPattern.SolidForeground;

                _s1P4Style.FillPattern = FillPattern.SolidForeground;
                _s2P4Style.FillPattern = FillPattern.SolidForeground;
                _s3P4Style.FillPattern = FillPattern.SolidForeground;
                _s4P4Style.FillPattern = FillPattern.SolidForeground;
                _s5P4Style.FillPattern = FillPattern.SolidForeground;
                _s6P4Style.FillPattern = FillPattern.SolidForeground;
                _s7P4Style.FillPattern = FillPattern.SolidForeground;
                _s8P4Style.FillPattern = FillPattern.SolidForeground;
                _s9P4Style.FillPattern = FillPattern.SolidForeground;
                _s10P4Style.FillPattern = FillPattern.SolidForeground;
                _s11P4Style.FillPattern = FillPattern.SolidForeground;
                _s12P4Style.FillPattern = FillPattern.SolidForeground;
                _s13P4Style.FillPattern = FillPattern.SolidForeground;
                _s14P4Style.FillPattern = FillPattern.SolidForeground;
                _s15P4Style.FillPattern = FillPattern.SolidForeground;
                _s16P4Style.FillPattern = FillPattern.SolidForeground;
                _s17P4Style.FillPattern = FillPattern.SolidForeground;
                _s18P4Style.FillPattern = FillPattern.SolidForeground;
                _s19P4Style.FillPattern = FillPattern.SolidForeground;
                _s20P4Style.FillPattern = FillPattern.SolidForeground;
                _s21P4Style.FillPattern = FillPattern.SolidForeground;
                _s22P4Style.FillPattern = FillPattern.SolidForeground;
                _s23P4Style.FillPattern = FillPattern.SolidForeground;
                _s24P4Style.FillPattern = FillPattern.SolidForeground;

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

                    #region P4

                    SetP4S1Value(row, index);
                    SetP4S2Value(row, index);
                    SetP4S3Value(row, index);
                    SetP4S4Value(row, index);
                    SetP4S5Value(row, index);
                    SetP4S6Value(row, index);
                    SetP4S7Value(row, index);
                    SetP4S8Value(row, index);
                    SetP4S9Value(row, index);
                    SetP4S10Value(row, index);
                    SetP4S11Value(row, index);
                    SetP4S12Value(row, index);
                    SetP4S13Value(row, index);
                    SetP4S14Value(row, index);
                    SetP4S15Value(row, index);
                    SetP4S16Value(row, index);
                    SetP4S17Value(row, index);
                    SetP4S18Value(row, index);
                    SetP4S19Value(row, index);
                    SetP4S20Value(row, index);
                    SetP4S21Value(row, index);
                    SetP4S22Value(row, index);
                    SetP4S23Value(row, index);
                    SetP4S24Value(row, index);
                    SetP4S25Value(row, index);

                    #endregion

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

            var s1p4 = row.CreateCell(StepLength.SourceLength * 9);
            var s2p4 = row.CreateCell(StepLength.SourceLength * 10);
            var s3p4 = row.CreateCell(StepLength.SourceLength * 11);
            var s4p4 = row.CreateCell(StepLength.SourceLength * 12);
            var s5p4 = row.CreateCell(StepLength.SourceLength * 13);
            var s6p4 = row.CreateCell(StepLength.SourceLength * 14);
            var s7p4 = row.CreateCell(StepLength.SourceLength * 15 + 3 * 0);
            var s8p4 = row.CreateCell(StepLength.SourceLength * 15 + 3 * 1);
            var s9p4 = row.CreateCell(StepLength.SourceLength * 15 + 3 * 2);
            var s10p4 = row.CreateCell(StepLength.SourceLength * 15 + 3 * 3);
            var s11p4 = row.CreateCell(StepLength.SourceLength * 15 + 3 * 4);
            var s12p4 = row.CreateCell(StepLength.SourceLength * 15 + 3 * 5);
            var s13p4 = row.CreateCell(StepLength.SourceLength * 15 + 3 * 6);
            var s14p4 = row.CreateCell(StepLength.SourceLength * 15 + 3 * 7);
            var s15p4 = row.CreateCell(StepLength.SourceLength * 15 + 3 * 8);
            var s16p4 = row.CreateCell(StepLength.SourceLength * 15 + 3 * 9);
            var s17p4 = row.CreateCell(StepLength.SourceLength * 15 + 3 * 10);
            var s18p4 = row.CreateCell(StepLength.SourceLength * 15 + 3 * 11);
            var s19p4 = row.CreateCell(StepLength.SourceLength * 21);
            var s20p4 = row.CreateCell(StepLength.SourceLength * 22);
            var s21p4 = row.CreateCell(StepLength.SourceLength * 23);
            var s22p4 = row.CreateCell(StepLength.SourceLength * 24);
            var s23p4 = row.CreateCell(StepLength.SourceLength * 25);
            var s24p4 = row.CreateCell(StepLength.SourceLength * 26);
            var s251p4 = row.CreateCell(StepLength.SourceLength * 27 + 0);
            var s252p4 = row.CreateCell(StepLength.SourceLength * 27 + 1);

            s1p4.SetCellValue("A总");
            s2p4.SetCellValue("B总");
            s3p4.SetCellValue("C总");
            s4p4.SetCellValue("1总");
            s5p4.SetCellValue("2总");
            s6p4.SetCellValue("3总");
            s7p4.SetCellValue("A三");
            s8p4.SetCellValue("B三");
            s9p4.SetCellValue("C三");
            s10p4.SetCellValue("1三");
            s11p4.SetCellValue("2三");
            s12p4.SetCellValue("3三");
            s13p4.SetCellValue("A热");
            s14p4.SetCellValue("B热");
            s15p4.SetCellValue("C热");
            s16p4.SetCellValue("1热");
            s17p4.SetCellValue("2热");
            s18p4.SetCellValue("3热");
            s19p4.SetCellValue("A序");
            s20p4.SetCellValue("B许");
            s21p4.SetCellValue("C序");
            s22p4.SetCellValue("1序");
            s23p4.SetCellValue("2序");
            s24p4.SetCellValue("3序");
            s251p4.SetCellValue("1");
            s252p4.SetCellValue("2");

            s1p4.CellStyle = _headerStyle;
            s2p4.CellStyle = _headerStyle;
            s3p4.CellStyle = _headerStyle;
            s4p4.CellStyle = _headerStyle;
            s5p4.CellStyle = _headerStyle;
            s6p4.CellStyle = _headerStyle;
            s7p4.CellStyle = _headerStyle;
            s8p4.CellStyle = _headerStyle;
            s9p4.CellStyle = _headerStyle;
            s10p4.CellStyle = _headerStyle;
            s11p4.CellStyle = _headerStyle;
            s12p4.CellStyle = _headerStyle;
            s13p4.CellStyle = _headerStyle;
            s14p4.CellStyle = _headerStyle;
            s15p4.CellStyle = _headerStyle;
            s16p4.CellStyle = _headerStyle;
            s17p4.CellStyle = _headerStyle;
            s18p4.CellStyle = _headerStyle;
            s19p4.CellStyle = _headerStyle;
            s20p4.CellStyle = _headerStyle;
            s21p4.CellStyle = _headerStyle;
            s22p4.CellStyle = _headerStyle;
            s23p4.CellStyle = _headerStyle;
            s24p4.CellStyle = _headerStyle;
            s251p4.CellStyle = _headerStyle;
            s252p4.CellStyle = _headerStyle;

            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 0, StepLength.SourceLength * 1 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 1, StepLength.SourceLength * 2 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 2, StepLength.SourceLength * 3 - 1));

            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 3, StepLength.SourceLength * 4 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 4, StepLength.SourceLength * 5 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 5, StepLength.SourceLength * 6 - 1));

            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 6, StepLength.SourceLength * 7 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 7, StepLength.SourceLength * 8 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 8, StepLength.SourceLength * 9 - 1));

            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 9, StepLength.SourceLength * 10 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 10, StepLength.SourceLength * 11 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 11, StepLength.SourceLength * 12 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 12, StepLength.SourceLength * 13 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 13, StepLength.SourceLength * 14 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 14, StepLength.SourceLength * 15 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 15 + 3 * 0, StepLength.SourceLength * 15 + 3 * 1 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 15 + 3 * 1, StepLength.SourceLength * 15 + 3 * 2 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 15 + 3 * 2, StepLength.SourceLength * 15 + 3 * 3 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 15 + 3 * 3, StepLength.SourceLength * 15 + 3 * 4 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 15 + 3 * 4, StepLength.SourceLength * 15 + 3 * 5 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 15 + 3 * 5, StepLength.SourceLength * 15 + 3 * 6 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 15 + 3 * 6, StepLength.SourceLength * 15 + 3 * 7 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 15 + 3 * 7, StepLength.SourceLength * 15 + 3 * 8 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 15 + 3 * 8, StepLength.SourceLength * 15 + 3 * 9 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 15 + 3 * 9, StepLength.SourceLength * 15 + 3 * 10 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 15 + 3 * 10, StepLength.SourceLength * 15 + 3 * 11 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 15 + 3 * 11, StepLength.SourceLength * 15 + 3 * 12 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 21, StepLength.SourceLength * 22 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 22, StepLength.SourceLength * 23 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 23, StepLength.SourceLength * 24 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 24, StepLength.SourceLength * 25 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 25, StepLength.SourceLength * 26 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 26, StepLength.SourceLength * 27 - 1));
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
                    cell.SetCellValue(i % StepLength.SourceLength);
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

        #region SetP4Value

        private void SetP4S1Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 9;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(1, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s1P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private void SetP4S2Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 10;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(2, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s2P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private void SetP4S3Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 11;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(3, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s3P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private void SetP4S4Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 12;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(4, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s4P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private void SetP4S5Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 13;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(5, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s5P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private void SetP4S6Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 14;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(6, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s6P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private void SetP4S7Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 0;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(7, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s7P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S8Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 1;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(8, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s8P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S9Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 2;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(9, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s9P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S10Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 3;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(10, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s10P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S11Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 4;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(11, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s11P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S12Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 5;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(12, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s12P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S13Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 6;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(13, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s13P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S14Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 7;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(14, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s14P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S15Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 8;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(15, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s15P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S16Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 9;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(16, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s16P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S17Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 10;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(17, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s17P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S18Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 11;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(18, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s18P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S19Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 21;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(19, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s19P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }
        private void SetP4S20Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 22;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(20, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s20P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }
        private void SetP4S21Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 23;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(21, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s21P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }
        private void SetP4S22Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 24;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(22, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s22P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }
        private void SetP4S23Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 25;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(23, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s23P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }
        private void SetP4S24Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 26;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP4Value<int>(24, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s24P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private void SetP4S25Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 27;
            var value = _context.GetP4Value<int>(25, rowIndex, 0);
            var cell = row.CreateCell(beforeColumn);
            cell.CellStyle = _s22P4Style;
            cell.SetCellValue(value);

            var value2 = _context.GetP4Value<int>(25, rowIndex, 1);
            var cell2 = row.CreateCell(beforeColumn+1);
            cell2.CellStyle = _s23P4Style;
            cell2.SetCellValue(value2);
        }

        #endregion


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