using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using HeProject.Model;

namespace HeProject
{
    public class WriteToExcel
    {
        private readonly ProcessContext _context;
        public WriteToExcel(ProcessContext context)
        {
            _context = context;
        }

        public void Write()
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");
            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            int x = 1;
            for (int i = 0; i < _context.Capacity; i++)
            {
                IRow row = sheet1.CreateRow(i);
                SetP1Value(row, i);
                SetP2Value(row, i);
                SetP3Value(row, i);
                SetP4Value(row, i);
                SetP5Value(row, i);
                SetP6Value(row, i);
                SetP7Value(row, i);
                SetP8Value(row, i);
            }

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }

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
    }
}
