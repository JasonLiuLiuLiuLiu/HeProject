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
        private readonly int _firstRow;
        private ICellStyle _headerStyle;
        public WriteToExcel(ProcessContext context)
        {
            _firstRow = 1;
            _context = context;
        }

        public void Write()
        {
            try
            {
                IWorkbook workbook = new XSSFWorkbook();
                _headerStyle = workbook.CreateCellStyle();
                ISheet sheet1 = workbook.CreateSheet("Sheet1");
                sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
                SetHeader(sheet1, 0);
                for (int i = _firstRow; i < _context.Capacity + _firstRow; i++)
                {
                    IRow row = sheet1.CreateRow(i);
                    var index = i - _firstRow;
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

        #region SetValue


        private void SetP1Value(IRow row, int rowIndex)
        {
            for (int i = 0; i < StepLength.P1; i++)
            {
                row.CreateCell(i).SetCellValue(_context.GetValue<int>(1, rowIndex, i));
            }
        }
        private void SetP2Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1;
            for (int i = beforeColumn; i < StepLength.P2 + beforeColumn; i++)
            {
                var value = _context.GetValue<bool>(2, rowIndex, i - beforeColumn);
                if (value)
                    row.CreateCell(i).SetCellValue("🔺");
            }
        }
        private void SetP3Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2;
            for (int i = beforeColumn; i < StepLength.P3 + beforeColumn; i++)
            {
                row.CreateCell(i).SetCellValue(_context.GetValue<int>(3, rowIndex, i - beforeColumn));
            }
        }

        private void SetP4Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3;
            for (int i = beforeColumn; i < StepLength.P4 + beforeColumn; i++)
            {
                var value = _context.GetValue<bool>(4, rowIndex, i - beforeColumn);
                if (value)
                    row.CreateCell(i).SetCellValue(i - beforeColumn);
            }
        }

        private void SetP5Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4;
            for (int i = beforeColumn; i < StepLength.P5 + beforeColumn; i++)
            {
                row.CreateCell(i).SetCellValue(_context.GetValue<int>(5, rowIndex, i - beforeColumn));
            }
        }

        private void SetP6Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5;
            for (int i = beforeColumn; i < StepLength.P6 + beforeColumn; i++)
            {
                var value = _context.GetValue<bool>(6, rowIndex, i - beforeColumn);
                if (value)
                    row.CreateCell(i).SetCellValue(i - beforeColumn);
            }
        }

        private void SetP7Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6;
            for (int i = beforeColumn; i < StepLength.P7 + beforeColumn; i++)
            {
                row.CreateCell(i).SetCellValue(_context.GetValue<int>(7, rowIndex, i - beforeColumn));
            }
        }
        private void SetP8Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 + StepLength.P7;
            for (int i = beforeColumn; i < StepLength.P8 + beforeColumn; i++)
            {
                row.CreateCell(i).SetCellValue(_context.GetValue<int>(8, rowIndex, i - beforeColumn));
            }
        }
        private void SetP9Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.P1 + StepLength.P2 + StepLength.P3 + StepLength.P4 + StepLength.P5 + StepLength.P6 + StepLength.P7 + StepLength.P8;
            for (int i = beforeColumn; i < StepLength.P9 + beforeColumn; i++)
            {
                row.CreateCell(i).SetCellValue(_context.GetValue<int>(9, rowIndex, i - beforeColumn));
            }
        }

        #endregion

    }
}
