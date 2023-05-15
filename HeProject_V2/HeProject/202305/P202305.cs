using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace HeProject._202305
{
    public class P202305
    {
        private IWorkbook _workBook;
        private Dictionary<int, bool?[]> _history;
        public P202305(IWorkbook workbook)
        {
            _workBook = workbook;
            _history = CalculateHistory();
        }
        public void Start()
        {
            SetHeader();
        }

        public void Display1()
        {
            foreach(var rowItem in _history)
            {
                var rowIndex = rowItem.Key;
                for (var i= 0; i < 9; i++)
                {

                }
            }
        }

        private KeyValuePair<int,bool>? PreValue(int row, int column)
        {
            if (row == 1)
                return null;

            if(!_history.ContainsKey(row))
                return null;

            if (_history[row][column].HasValue)
                return new KeyValuePair<int, bool>(row, _history[row][column].Value);

            return PreValue(row - 1, column);
        }

        public void SetHeader()
        {
            var sheet = _workBook.GetSheetAt(0);
            var headerRow = sheet.GetRow(0);
            for (int i = 0; i < 21; i++)
            {
                var cell = headerRow.CreateCell(i + 278);
                cell.SetCellValue(i + 1);
                sheet.SetColumnWidth(i + 278, 700);
            }
        }

        private Dictionary<int, bool?[]> CalculateHistory()
        {
            var history = new Dictionary<int, bool?[]>();
            var sheet = _workBook.GetSheetAt(0);
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                var rowResult = new bool?[9];
                history.Add(i, rowResult);
                var row = sheet.GetRow(i);
                if (row == null)
                    continue;
                for (int c = 269; c <= 277; c++)
                {
                    var cell = row.GetCell(c);
                    if (cell == null) continue;
                    var rowIndex = c - 269;
                    if (cell.CellStyle.FillForegroundColor == 10)
                    {
                        rowResult[rowIndex] = true;
                    }
                    else if (cell.CellStyle.FillForegroundColor == 12)
                    {
                        rowResult[rowIndex] = false;
                    }
                    else
                    {
                        rowResult[rowIndex] = null;
                    }
                }
            }
            return history;
        }
    }
}
