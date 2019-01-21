using System;
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
        private readonly int _p1FirstRow;
        private readonly int _p2FirstRow;
        private ICellStyle _headerStyle;
        private ICellStyle _s1Style;
        private ICellStyle _s2Style;
        private ICellStyle _s3Style;
        private ICellStyle _s4Style;
        private ICellStyle _s5Style;
        private ICellStyle _s6Style;
        private ICellStyle _s7Style;
        private ICellStyle _s8Style;
        private ICellStyle _s9Style;
        public WriteToExcel(ProcessContext context)
        {
            _p1FirstRow = 1;
            _p2FirstRow = _p1FirstRow + context.Capacity + 2;
            _context = context;
        }

        public void Write()
        {
            try
            {
                IWorkbook workbook = new XSSFWorkbook();
                _headerStyle = workbook.CreateCellStyle();

                #region SetStyle

                _s1Style = workbook.CreateCellStyle();
                _s2Style = workbook.CreateCellStyle();
                _s3Style = workbook.CreateCellStyle();
                _s4Style = workbook.CreateCellStyle();
                _s5Style = workbook.CreateCellStyle();
                _s6Style = workbook.CreateCellStyle();
                _s7Style = workbook.CreateCellStyle();
                _s8Style = workbook.CreateCellStyle();
                _s9Style = workbook.CreateCellStyle();
                _s1Style.BottomBorderColor = IndexedColors.Black.Index;
                _s1Style.LeftBorderColor = IndexedColors.Black.Index;
                _s1Style.RightBorderColor = IndexedColors.Black.Index;
                _s1Style.LeftBorderColor = IndexedColors.Black.Index;

                _s2Style.BottomBorderColor = IndexedColors.Black.Index;
                _s2Style.LeftBorderColor = IndexedColors.Black.Index;
                _s2Style.RightBorderColor = IndexedColors.Black.Index;
                _s2Style.LeftBorderColor = IndexedColors.Black.Index;

                _s3Style.BottomBorderColor = IndexedColors.Black.Index;
                _s3Style.LeftBorderColor = IndexedColors.Black.Index;
                _s3Style.RightBorderColor = IndexedColors.Black.Index;
                _s3Style.LeftBorderColor = IndexedColors.Black.Index;

                _s4Style.BottomBorderColor = IndexedColors.Black.Index;
                _s4Style.LeftBorderColor = IndexedColors.Black.Index;
                _s4Style.RightBorderColor = IndexedColors.Black.Index;
                _s4Style.LeftBorderColor = IndexedColors.Black.Index;

                _s5Style.BottomBorderColor = IndexedColors.Black.Index;
                _s5Style.LeftBorderColor = IndexedColors.Black.Index;
                _s5Style.RightBorderColor = IndexedColors.Black.Index;
                _s5Style.LeftBorderColor = IndexedColors.Black.Index;

                _s6Style.BottomBorderColor = IndexedColors.Black.Index;
                _s6Style.LeftBorderColor = IndexedColors.Black.Index;
                _s6Style.RightBorderColor = IndexedColors.Black.Index;
                _s6Style.LeftBorderColor = IndexedColors.Black.Index;

                _s7Style.BottomBorderColor = IndexedColors.Black.Index;
                _s7Style.LeftBorderColor = IndexedColors.Black.Index;
                _s7Style.RightBorderColor = IndexedColors.Black.Index;
                _s7Style.LeftBorderColor = IndexedColors.Black.Index;

                _s8Style.BottomBorderColor = IndexedColors.Black.Index;
                _s8Style.LeftBorderColor = IndexedColors.Black.Index;
                _s8Style.RightBorderColor = IndexedColors.Black.Index;
                _s8Style.LeftBorderColor = IndexedColors.Black.Index;

                _s9Style.BottomBorderColor = IndexedColors.Black.Index;
                _s9Style.LeftBorderColor = IndexedColors.Black.Index;
                _s9Style.RightBorderColor = IndexedColors.Black.Index;
                _s9Style.LeftBorderColor = IndexedColors.Black.Index;

                _s1Style.FillBackgroundColor = IndexedColors.Aqua.Index;
                _s2Style.FillBackgroundColor = IndexedColors.Blue.Index;
                _s3Style.FillBackgroundColor = IndexedColors.BlueGrey.Index;
                _s4Style.FillBackgroundColor = IndexedColors.Brown.Index;
                _s5Style.FillBackgroundColor = IndexedColors.CornflowerBlue.Index;
                _s6Style.FillBackgroundColor = IndexedColors.Coral.Index;
                _s7Style.FillBackgroundColor = IndexedColors.DarkGreen.Index;
                _s8Style.FillBackgroundColor = IndexedColors.Yellow.Index;
                _s9Style.FillBackgroundColor = IndexedColors.Violet.Index;

                #endregion


                ISheet sheet1 = workbook.CreateSheet("Sheet1");
                sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
                SetHeader(sheet1, 0);
                for (int i = _p1FirstRow; i < _context.Capacity + _p1FirstRow; i++)
                {
                    IRow row = sheet1.CreateRow(i);
                    var index = i - _p1FirstRow;
                    SetP1Value(row, index);
                    SetP2Value(row, index);
                    SetP3Value(row, index);
                    SetP4Value(row, index);
                    SetP5Value(row, index);
                    SetP6Value(row, index);
                    SetP7Value(row, index);
                    SetP8Value(row, index);
                    SetP9Value(row, index);
                }

                for (int i = _p2FirstRow; i < _context.Capacity + _p2FirstRow; i++)
                {
                    IRow row = sheet1.CreateRow(i);
                    var index = i - _p2FirstRow;
                    SetP1Value(row, index);
                    Set2P2Value(row, index);
                    Set2P3Value(row, index);
                    Set2P4Value(row, index);
                    Set2P5Value(row, index);
                    Set2P6Value(row, index);
                    Set2P7Value(row, index);
                    Set2P8Value(row, index);
                    Set2P9Value(row, index);
                }

                FileStream sw = File.Create("_Output.xlsx");
                workbook.Write(sw);
                sw.Close();
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
            var cellHeader1 = row.CreateCell(0, CellType.String);
            cellHeader1.SetCellValue("第一步");
            var cellHeader2 = row.CreateCell(StepLength.P1, CellType.String);
            cellHeader2.SetCellValue("第二步");
            var cellHeader3 = row.CreateCell(StepLength.P1 + StepLength.P2, CellType.String);
            cellHeader3.SetCellValue("第三步");
            var cellHeader4 = row.CreateCell(StepLength.P1 + StepLength.P2 + StepLength.P3, CellType.String);
            cellHeader4.SetCellValue("第四步");
            var cellHeader5 = row.CreateCell(StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4, CellType.String);
            cellHeader5.SetCellValue("第五步");
            var cellHeader6 = row.CreateCell(StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5, CellType.String);
            cellHeader6.SetCellValue("第六步");
            var cellHeader7 = row.CreateCell(StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6, CellType.String);
            cellHeader7.SetCellValue("第七步");
            var cellHeader8 = row.CreateCell(StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 + StepLength.P7, CellType.String);
            cellHeader8.SetCellValue("第八步");
            var cellHeader9 = row.CreateCell(StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 + StepLength.P7 + StepLength.P8, CellType.String);
            cellHeader9.SetCellValue("第九步");
            cellHeader1.CellStyle = _headerStyle;
            cellHeader2.CellStyle = _headerStyle;
            cellHeader3.CellStyle = _headerStyle;
            cellHeader4.CellStyle = _headerStyle;
            cellHeader5.CellStyle = _headerStyle;
            cellHeader6.CellStyle = _headerStyle;
            cellHeader7.CellStyle = _headerStyle;
            cellHeader8.CellStyle = _headerStyle;
            cellHeader9.CellStyle = _headerStyle;
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 0, StepLength.P1 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.P1, StepLength.P1 + StepLength.P2 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.P1 + StepLength.P2, StepLength.P1 + StepLength.P2 + StepLength.P3 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.P1 + StepLength.P2 + StepLength.P3, StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4, StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5, StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6, StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 + StepLength.P7 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 + StepLength.P7, StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 + StepLength.P7 + StepLength.P8 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 + StepLength.P7 + StepLength.P8, StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 + StepLength.P7 + StepLength.P8 + StepLength.P9 - 1));
        }

        #region SetP1Value


        private void SetP1Value(IRow row, int rowIndex)
        {
            for (int i = 0; i < StepLength.P1; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s1Style;
                cell.SetCellValue(_context.GetP1Value<int>(1, rowIndex, i));
            }
        }
        private void SetP2Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1;
            for (int i = beforeColumn; i < StepLength.P2 + beforeColumn; i++)
            {
                var value = _context.GetP1Value<bool>(2, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s2Style;
                if (value)
                    cell.SetCellValue("🔺");
            }
        }
        private void SetP3Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2;
            for (int i = beforeColumn; i < StepLength.P3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(_context.GetP1Value<int>(3, rowIndex, i - beforeColumn));
                cell.CellStyle = _s3Style;
            }
        }

        private void SetP4Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3;
            for (int i = beforeColumn; i < StepLength.P4 + beforeColumn; i++)
            {
                var value = _context.GetP1Value<bool>(4, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s4Style;
                if (value)
                    cell.SetCellValue(i - beforeColumn);
            }
        }

        private void SetP5Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4;
            for (int i = beforeColumn; i < StepLength.P5 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(_context.GetP1Value<int>(5, rowIndex, i - beforeColumn));
                cell.CellStyle = _s5Style;
            }
        }

        private void SetP6Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5;
            for (int i = beforeColumn; i < StepLength.P6 + beforeColumn; i++)
            {
                var value = _context.GetP1Value<bool>(6, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s6Style;
                if (value)
                    cell.SetCellValue(i - beforeColumn);
            }
        }

        private void SetP7Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6;
            for (int i = beforeColumn; i < StepLength.P7 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(_context.GetP1Value<int>(7, rowIndex, i - beforeColumn));
                cell.CellStyle = _s7Style;
            }
        }
        private void SetP8Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 + StepLength.P7;
            for (int i = beforeColumn; i < StepLength.P8 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s8Style;
                cell.SetCellValue(_context.GetP1Value<int>(8, rowIndex, i - beforeColumn));
            }
        }
        private void SetP9Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 + StepLength.P7 + StepLength.P8;
            for (int i = beforeColumn; i < StepLength.P9 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(_context.GetP1Value<int>(9, rowIndex, i - beforeColumn));
                cell.CellStyle = _s9Style;
            }
        }

        #endregion

        #region SetP2Value

       
        private void Set2P2Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1;
            for (int i = beforeColumn; i < StepLength.P2 + beforeColumn; i++)
            {
                var value = _context.GetP2Value<bool>(2, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s2Style;
                if (value)
                    cell.SetCellValue("⭕");
            }
        }
        private void Set2P3Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2;
            for (int i = beforeColumn; i < StepLength.P3 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(_context.GetP2Value<int>(3, rowIndex, i - beforeColumn));
                cell.CellStyle = _s3Style;
            }
        }

        private void Set2P4Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3;
            for (int i = beforeColumn; i < StepLength.P4 + beforeColumn; i++)
            {
                var value = _context.GetP2Value<bool>(4, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s4Style;
                if (value)
                    cell.SetCellValue(i - beforeColumn);
            }
        }

        private void Set2P5Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4;
            for (int i = beforeColumn; i < StepLength.P5 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(_context.GetP2Value<int>(5, rowIndex, i - beforeColumn));
                cell.CellStyle = _s5Style;
            }
        }

        private void Set2P6Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5;
            for (int i = beforeColumn; i < StepLength.P6 + beforeColumn; i++)
            {
                var value = _context.GetP2Value<bool>(6, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s6Style;
                if (value)
                    cell.SetCellValue(i - beforeColumn);
            }
        }

        private void Set2P7Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6;
            for (int i = beforeColumn; i < StepLength.P7 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(_context.GetP2Value<int>(7, rowIndex, i - beforeColumn));
                cell.CellStyle = _s7Style;
            }
        }
        private void Set2P8Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 + StepLength.P7;
            for (int i = beforeColumn; i < StepLength.P8 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = _s8Style;
                cell.SetCellValue(_context.GetP2Value<int>(8, rowIndex, i - beforeColumn));
            }
        }
        private void Set2P9Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 + StepLength.P7 + StepLength.P8;
            for (int i = beforeColumn; i < StepLength.P9 + beforeColumn; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(_context.GetP2Value<int>(9, rowIndex, i - beforeColumn));
                cell.CellStyle = _s9Style;
            }
        }

        #endregion

    }
}
