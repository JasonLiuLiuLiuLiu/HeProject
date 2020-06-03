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
        private Dictionary<int, ProcessContext> _resultDic = new Dictionary<int, ProcessContext>(6);
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

        private ICellStyle _blueStyle;
        private ICellStyle _yellowStyle;

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

                _blueStyle = workbook.CreateCellStyle();
                _yellowStyle = workbook.CreateCellStyle();

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

                _blueStyle.FillForegroundColor = 12;
                _yellowStyle.FillForegroundColor = 43;

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

                _blueStyle.FillPattern = FillPattern.SolidForeground;
                _yellowStyle.FillPattern = FillPattern.SolidForeground;

                #endregion SetStyle

                ISheet sheet1 = workbook.CreateSheet("Sheet1");
                sheet1.DefaultColumnWidth = 1;
                sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
                SetHeader(sheet1, 0);
                for (int i = 1; i <= _context.Capacity; i++)
                {
                    IRow row = sheet1.CreateRow(i);
                    if (i == _context.Capacity)
                    {
                        row = sheet1.CreateRow(i + 1);
                    }
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
                    SetP4S26Value(row, index);
                    SetP4S27Value(row, index);
                    SetP4S28Value(row, index);
                    SetP4S30Value(row, index);
                    SetP4S31Value(row, index);
                    SetP4S32Value(row, index);
                    SetP4S33Value(row, index);
                    SetP4S34Value(row, index);

                    #endregion

                }

                for (int i = 1; i < 6; i++)
                {
                    IRow row = sheet1.CreateRow(_context.Capacity + 1 + i * 2);

                    var context = _resultDic[i];

                    var index = context.Capacity - 1;

                    #region P1

                    SetP1S1Value(row, index, context);
                    SetP1S2Value(row, index, context);
                    SetP1S3Value(row, index, context);

                    #endregion P1

                    #region P2

                    SetP2S1Value(row, index, context);
                    SetP2S2Value(row, index, context);
                    SetP2S3Value(row, index, context);

                    #endregion P2

                    #region P3

                    SetP3S1Value(row, index, context);
                    SetP3S2Value(row, index, context);
                    SetP3S3Value(row, index, context);

                    #endregion P3

                    #region P4

                    SetP4S1Value(row, index, context);
                    SetP4S2Value(row, index, context);
                    SetP4S3Value(row, index, context);
                    SetP4S4Value(row, index, context);
                    SetP4S5Value(row, index, context);
                    SetP4S6Value(row, index, context);
                    SetP4S7Value(row, index, context);
                    SetP4S8Value(row, index, context);
                    SetP4S9Value(row, index, context);
                    SetP4S10Value(row, index, context);
                    SetP4S11Value(row, index, context);
                    SetP4S12Value(row, index, context);
                    SetP4S13Value(row, index, context);
                    SetP4S14Value(row, index, context);
                    SetP4S15Value(row, index, context);
                    SetP4S16Value(row, index, context);
                    SetP4S17Value(row, index, context);
                    SetP4S18Value(row, index, context);
                    SetP4S19Value(row, index, context);
                    SetP4S20Value(row, index, context);
                    SetP4S21Value(row, index, context);
                    SetP4S22Value(row, index, context);
                    SetP4S23Value(row, index, context);
                    SetP4S24Value(row, index, context);
                    SetP4S25Value(row, index, context);
                    SetP4S26Value(row, index, context);
                    SetP4S27Value(row, index, context);
                    SetP4S28Value(row, index, context);
                    SetP4S30Value(row, index, context);
                    SetP4S31Value(row, index, context);
                    SetP4S32Value(row, index, context);
                    SetP4S33Value(row, index, context);
                    SetP4S34Value(row, index, context);
                    #endregion
                }
                P4S27Finally(sheet1);
                P4S32Finally(sheet1);

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
            var s26p4 = row.CreateCell(StepLength.SourceLength * 27 + 2);
            var s27p4 = row.CreateCell(StepLength.SourceLength * 28 + 2);
            var s28p4 = row.CreateCell(StepLength.SourceLength * 29 + 2);
            var s30p4 = row.CreateCell(StepLength.SourceLength * 29 + 6);
            var s31p4 = row.CreateCell(StepLength.SourceLength * 30 + 6);
            var s32p4 = row.CreateCell(StepLength.SourceLength * 31 + 6);
            var s33p4 = row.CreateCell(StepLength.SourceLength * 31 + 9);
            var s34p4 = row.CreateCell(StepLength.SourceLength * 31 + 12);

            var s27p4_v2 = row.CreateCell(StepLength.SourceLength * 31 + 15);
            var s27p4_v3 = row.CreateCell(StepLength.SourceLength * 31 + 21);
            var s27p4_v4 = row.CreateCell(StepLength.SourceLength * 31 + 22);
            var s27p4_v5 = row.CreateCell(StepLength.SourceLength * 31 + 28);

            var s32p4_v2 = row.CreateCell(StepLength.SourceLength * 31 + 29);
            var s32p4_v3 = row.CreateCell(StepLength.SourceLength * 31 + 33);
            var s32p4_v4 = row.CreateCell(StepLength.SourceLength * 31 + 34);
            var s32p4_v5 = row.CreateCell(StepLength.SourceLength * 31 + 38);

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
            s20p4.SetCellValue("B序");
            s21p4.SetCellValue("C序");
            s22p4.SetCellValue("1序");
            s23p4.SetCellValue("2序");
            s24p4.SetCellValue("3序");
            s251p4.SetCellValue("1");
            s252p4.SetCellValue("2");
            s26p4.SetCellValue("增1");
            s27p4.SetCellValue("增2");
            s28p4.SetCellValue("跳");
            s30p4.SetCellValue("增2序1");
            s31p4.SetCellValue("增2序2");
            s32p4.SetCellValue("总热");
            s33p4.SetCellValue("序热");
            s34p4.SetCellValue("增热");

            s27p4_v2.SetCellValue("变增2");
            s27p4_v3.SetCellValue("黄");
            s27p4_v4.SetCellValue("新增2");
            s27p4_v5.SetCellValue("蓝");
            s32p4_v2.SetCellValue("总序");
            s32p4_v3.SetCellValue("黄");
            s32p4_v4.SetCellValue("新总序");
            s32p4_v5.SetCellValue("蓝");

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
            s26p4.CellStyle = _headerStyle;
            s27p4.CellStyle = _headerStyle;
            s28p4.CellStyle = _headerStyle;
            s30p4.CellStyle = _headerStyle;
            s31p4.CellStyle = _headerStyle;
            s32p4.CellStyle = _headerStyle;
            s33p4.CellStyle = _headerStyle;
            s34p4.CellStyle = _headerStyle;
            s27p4_v2.CellStyle = _headerStyle;
            s27p4_v3.CellStyle = _headerStyle;
            s27p4_v4.CellStyle = _headerStyle;
            s27p4_v5.CellStyle = _headerStyle;
            s32p4_v2.CellStyle = _headerStyle;
            s32p4_v3.CellStyle = _headerStyle;
            s32p4_v4.CellStyle = _headerStyle;
            s32p4_v5.CellStyle = _headerStyle;

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
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 27 + 2, StepLength.SourceLength * 28 + 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 28 + 2, StepLength.SourceLength * 29 + 1));

            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 29 + 2, StepLength.SourceLength * 29 + 5));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 29 + 6, StepLength.SourceLength * 30 + 5));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 30 + 6, StepLength.SourceLength * 31 + 5));

            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 31 + 6, StepLength.SourceLength * 31 + 8));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 31 + 9, StepLength.SourceLength * 31 + 11));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 31 + 12, StepLength.SourceLength * 31 + 14));

            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 31 + 15, StepLength.SourceLength * 31 + 20));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 31 + 22, StepLength.SourceLength * 31 + 27));

            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 31 + 29, StepLength.SourceLength * 31 + 32));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 31 + 34, StepLength.SourceLength * 31 + 37));


            for (int i = StepLength.SourceLength * 29 + 2; i < StepLength.SourceLength * 29 + 6; i++)
            {
                sheet.SetColumnWidth(i, 1000);
            }

            for (int i = StepLength.SourceLength * 27; i < StepLength.SourceLength * 27 + 14; i++)
            {
                sheet.SetColumnWidth(i, 1000);
            }
        }

        #region SetP1Value



        private void SetP1S1Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = 0;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP1Value<bool>(1, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s1P1Style;
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
            int beforeColumn = StepLength.SourceLength * 1;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP1Value<bool>(2, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s2P1Style;
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
            int beforeColumn = StepLength.SourceLength * 2;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP1Value<bool>(3, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s3P1Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        #endregion SetP1Value

        #region SetP2Value

        private void SetP2S1Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 3;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP2Value<bool>(1, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s1P2Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP2S2Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 4;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP2Value<bool>(2, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s2P2Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP2S3Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 5;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP2Value<bool>(3, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s3P2Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        #endregion SetP2Value

        #region SetP3Value

        private void SetP3S1Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 6;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP3Value<bool>(1, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s1P3Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP3S2Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 7;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP3Value<bool>(2, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s2P3Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP3S3Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 8;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP3Value<bool>(3, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s3P3Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        #endregion SetP3Value

        #region SetP4Value

        private void SetP4S1Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 9;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(1, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s1P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private void SetP4S2Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 10;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(2, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s2P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private void SetP4S3Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 11;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(3, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s3P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private void SetP4S4Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 12;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(4, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s4P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private void SetP4S5Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 13;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(5, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s5P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private void SetP4S6Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 14;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(6, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s6P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private void SetP4S7Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 0;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(7, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s7P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S8Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 1;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(8, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s8P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S9Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 2;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(9, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s9P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S10Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 3;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(10, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s10P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S11Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 4;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(11, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s11P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S12Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 5;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(12, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s12P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S13Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 6;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(13, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s13P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S14Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 7;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(14, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s14P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S15Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 8;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(15, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s15P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S16Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 9;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(16, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s16P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S17Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 10;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(17, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s17P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S18Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 15 + 3 * 11;
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(18, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s18P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S19Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 21;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(19, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s19P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }
        private void SetP4S20Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 22;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(20, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s20P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }
        private void SetP4S21Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 23;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(21, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s21P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }
        private void SetP4S22Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 24;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(22, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s22P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }
        private void SetP4S23Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 25;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(23, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s23P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }
        private void SetP4S24Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 26;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(24, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s24P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private static int[] P4S25Blue = new[] { 5, 6, 7 };
        private void SetP4S25Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 27;
            var value = context.GetP4Value<int>(25, rowIndex, 0);
            var cell = row.CreateCell(beforeColumn);
            cell.CellStyle = P4S25Blue.Contains(value) ? _blueStyle : _s22P4Style;
            cell.SetCellValue(value);

            var value2 = context.GetP4Value<int>(25, rowIndex, 1);
            var cell2 = row.CreateCell(beforeColumn + 1);
            cell2.CellStyle = _s23P4Style;
            cell2.SetCellValue(value2);
        }

        private static int[] P4S26Blue = new[] { 2, 3, 4 };
        private void SetP4S26Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 27 + 2;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(26, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = P4S26Blue.Contains(value) ? _blueStyle : _s24P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }
        }

        private static int[] P4S27Yellow = new[] { 2, 3, 4 };
        private static int[][] P4S27FinallyRowPossible = new int[6][];
        private static int P4S27FinallyRowIndex = -1;
        private void SetP4S27Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }

            var finallyRow = rowIndex == (context.Capacity - 1);
            if (finallyRow)
            {
                P4S27FinallyRowIndex++;
                P4S27FinallyRowPossible[P4S27FinallyRowIndex] = new int[6];
            }

            int beforeColumn = StepLength.SourceLength * 28 + 2;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(27, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s23P4Style;
                if (value > 0)
                    cell.SetCellValue(value);
            }

            beforeColumn = StepLength.SourceLength * 31 + 15;
            var yellowCount = 0;
            var index = -1;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                index++;
                var value = context.GetP4Value<int>(27, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                if (P4S27Yellow.Contains(value))
                {
                    yellowCount++;
                    cell.CellStyle = _yellowStyle;
                }
                else
                {
                    cell.CellStyle = _s23P4Style;
                }

                if (finallyRow)
                {
                    P4S27FinallyRowPossible[P4S27FinallyRowIndex][index] = value;
                }
                if (value > 0)
                    cell.SetCellValue(value);
            }

            var cellCount = row.CreateCell(StepLength.SourceLength * 31 + 21);
            cellCount.CellStyle = _s24P4Style;
            cellCount.SetCellValue(yellowCount);

        }

        private static int[] P4S28Blue = new[] { 7, 8, 9 };
        private void SetP4S28Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 29 + 2;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = context.GetP4Value<int>(28, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = P4S28Blue.Contains(value) ? _blueStyle : _s22P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S30Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 29 + 6;
            var valueDic = context.GetP4RowResult(30, rowIndex);
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s23P4Style;
                if (valueDic.ContainsKey(i - beforeColumn) && (bool)valueDic[i - beforeColumn])
                    cell.SetCellValue(i - beforeColumn);
            }
        }

        private void SetP4S31Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 30 + 6;
            var valueDic = context.GetP4RowResult(31, rowIndex);
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s24P4Style;
                if (valueDic.ContainsKey(i - beforeColumn) && (bool)valueDic[i - beforeColumn])
                    cell.SetCellValue(i - beforeColumn);
            }
        }

        private static int[] P4S32Yellow = new[] { 6, 7, 8 };
        private static Dictionary<int, int> P4S32YellowDic = new Dictionary<int, int>();
        private static int[][] P4S32FinallyRowPossible = new int[6][];
        private static int P4S32FinallyRowIndex = -1;
        private void SetP4S32Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            var finallyRow = rowIndex == (context.Capacity - 1);
            if (finallyRow)
            {
                P4S32FinallyRowIndex++;
                P4S32FinallyRowPossible[P4S32FinallyRowIndex] = new int[4];
            }
            int beforeColumn = StepLength.SourceLength * 31 + 6;
            var valueDic = rowIndex > 4 ? context.GetP4RowResult(32, rowIndex) : new Dictionary<int, object>() { { 0, 0 }, { 1, 0 }, { 2, 0 } };
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s22P4Style;
                cell.SetCellValue((int)valueDic[i - beforeColumn]);
            }
            for (int i = beforeColumn; i < 2 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i + 23);
                var value = (int)valueDic[i - beforeColumn];
                var exist = P4S32Yellow.Contains(value);
                if (exist)
                {
                    if (P4S32YellowDic.ContainsKey(rowIndex))
                    {
                        P4S32YellowDic[rowIndex] = P4S32YellowDic[rowIndex] + 1;
                    }
                    else
                    {
                        P4S32YellowDic.Add(rowIndex, 1);
                    }
                }
                if (finallyRow)
                {
                    P4S32FinallyRowPossible[P4S32FinallyRowIndex][i - beforeColumn] = value;
                }
                cell.CellStyle = exist ? _yellowStyle : _s23P4Style;
                cell.SetCellValue(value);
            }
        }

        private void SetP4S33Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            var finallyRow = rowIndex == (context.Capacity - 1);
            int beforeColumn = StepLength.SourceLength * 31 + 9;
            var valueDic = rowIndex > 4 ? context.GetP4RowResult(33, rowIndex) : new Dictionary<int, object>() { { 0, 0 }, { 1, 0 }, { 2, 0 } };
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s23P4Style;
                cell.SetCellValue((int)valueDic[i - beforeColumn]);
            }
            for (int i = beforeColumn; i < 2 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i + 22);
                var value = (int)valueDic[i - beforeColumn];
                var exist = P4S32Yellow.Contains(value);
                if (exist)
                {
                    if (P4S32YellowDic.ContainsKey(rowIndex))
                    {
                        P4S32YellowDic[rowIndex] = P4S32YellowDic[rowIndex] + 1;
                    }
                    else
                    {
                        P4S32YellowDic.Add(rowIndex, 1);
                    }
                }
                if (finallyRow)
                {
                    P4S32FinallyRowPossible[P4S32FinallyRowIndex][i - beforeColumn + 2] = value;
                }
                cell.CellStyle = exist ? _yellowStyle : _s23P4Style;
                cell.SetCellValue(value);
            }
            var cellCount = row.CreateCell(2 + beforeColumn + 22);
            cellCount.CellStyle = _s24P4Style;
            cellCount.SetCellValue(P4S32YellowDic.ContainsKey(rowIndex) ? P4S32YellowDic[rowIndex] : 0);
        }

        private void SetP4S34Value(IRow row, int rowIndex, ProcessContext context = null)
        {
            if (context == null)
            {
                context = _context;
            }
            int beforeColumn = StepLength.SourceLength * 31 + 12;
            var valueDic = rowIndex > 4 ? context.GetP4RowResult(34, rowIndex) : new Dictionary<int, object>() { { 0, 0 }, { 1, 0 }, { 2, 0 } };
            for (int i = beforeColumn; i < 3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s24P4Style;
                cell.SetCellValue((int)valueDic[i - beforeColumn]);
            }
        }

        #endregion

        private void P4S27Finally(ISheet sheet)
        {
            bool[][] result = new bool[6][];
            for (int i = 0; i < 6; i++)
            {
                result[i] = new bool[6];
            }
            for (int i = 0; i < 6; i++)
            {
                var yellowIndex = new List<int>();
                for (int j = 0; j < 6; j++)
                {
                    if (P4S27Yellow.Contains(P4S27FinallyRowPossible[j][i]))
                    {
                        yellowIndex.Add(j);
                    }
                }

                if (yellowIndex.Count == 1 || yellowIndex.Count == 2)
                {
                    for (int j = 0; j < 6; j++)
                    {
                        if (!yellowIndex.Contains(j))
                        {
                            result[j][i] = true;
                        }
                    }
                }
                else if (yellowIndex.Count == 4 || yellowIndex.Count == 5)
                {
                    for (int j = 0; j < 6; j++)
                    {
                        if (yellowIndex.Contains(j))
                        {
                            result[j][i] = true;
                        }
                    }
                }
            }

            var rowIndex = -1;
            for (int i = 0; i < 6; i++)
            {
                rowIndex++;
                IRow row = sheet.GetRow(_context.Capacity + 1 + i * 2);
                for (int j = 0; j < 6; j++)
                {
                    var cell = row.CreateCell(StepLength.SourceLength * 31 + 22 + j);
                    cell.SetCellValue(P4S27FinallyRowPossible[rowIndex][j]);
                    cell.CellStyle = result[rowIndex][j] ? _blueStyle : _s23P4Style;
                }
                var cellCount = row.CreateCell(StepLength.SourceLength * 31 + 28);
                cellCount.CellStyle = _s24P4Style;
                cellCount.SetCellValue(result[rowIndex].Count(u => u));
            }
        }

        private void P4S32Finally(ISheet sheet)
        {
            bool[][] result = new bool[6][];
            for (int i = 0; i < 6; i++)
            {
                result[i] = new bool[4];
            }
            for (int i = 0; i < 4; i++)
            {
                var yellowIndex = new List<int>();
                for (int j = 0; j < 6; j++)
                {
                    if (P4S32Yellow.Contains(P4S32FinallyRowPossible[j][i]))
                    {
                        yellowIndex.Add(j);
                    }
                }

                if (yellowIndex.Count == 1 || yellowIndex.Count == 2)
                {
                    for (int j = 0; j < 6; j++)
                    {
                        if (!yellowIndex.Contains(j))
                        {
                            result[j][i] = true;
                        }
                    }
                }
                else if (yellowIndex.Count == 4 || yellowIndex.Count == 5)
                {
                    for (int j = 0; j < 6; j++)
                    {
                        if (yellowIndex.Contains(j))
                        {
                            result[j][i] = true;
                        }
                    }
                }
            }

            var rowIndex = -1;
            for (int i = 0; i < 6; i++)
            {
                rowIndex++;
                IRow row = sheet.GetRow(_context.Capacity + 1 + i * 2);
                for (int j = 0; j < 4; j++)
                {
                    var cell = row.CreateCell(StepLength.SourceLength * 31 + 34 + j);
                    cell.SetCellValue(P4S32FinallyRowPossible[rowIndex][j]);
                    cell.CellStyle = result[rowIndex][j] ? _blueStyle : _s23P4Style;
                }
                var cellCount = row.CreateCell(StepLength.SourceLength * 31 + 38);
                cellCount.CellStyle = _s24P4Style;
                cellCount.SetCellValue(result[rowIndex].Count(u => u));
            }
        }

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