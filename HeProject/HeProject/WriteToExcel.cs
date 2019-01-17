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
                    row.CreateCell(i).SetCellValue(value);
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
    }
}
