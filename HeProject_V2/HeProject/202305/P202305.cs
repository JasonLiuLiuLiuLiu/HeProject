using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HeProject._202305
{
    public class P202305
    {
        private IWorkbook _workBook;
        private ISheet _sheet;
        private int _lastRowIndex;
        private Dictionary<int, bool?[]> _history;
        public P202305(IWorkbook workbook)
        {
            _workBook = workbook;
            _sheet = _workBook.GetSheetAt(0);
            _lastRowIndex = _sheet.LastRowNum - 1 - 12;
            _history = CalculateHistory();
        }
        public void Start()
        {
            SetHeader();
            Display1();
            Display2();
            Display3();
            Display4();
            Display5();
            Display6();
            Display7();
            Display8();
            Display9();
            Display10();
            Display11();
            Display12();
            Display13();
            Display14();
        }

        private Dictionary<int, List<KeyValuePair<int, bool?>>> _display1 = new Dictionary<int, List<KeyValuePair<int, bool?>>>();
        public void Display1()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;

                var value = 0;
                var hitColumns = new List<KeyValuePair<int, bool?>>();
                rowIndex = GetPreRow(rowIndex);
                if (_history.ContainsKey(rowIndex))
                {
                    for (var i = 0; i < 9; i++)
                    {
                        var value1 = PreValue(rowIndex, i);
                        if (value1.HasValue)
                        {
                            var value2 = PreValue(GetPreRow(value1.Value.Key), i);
                            if (value1.Value.Value == value2?.Value)
                            {
                                hitColumns.Add(new KeyValuePair<int, bool?>(i, value1.Value.Value));
                                value++;
                            }

                        }
                    }
                }

                var cell = _sheet.GetRow(rowItem.Key)?.CreateCell(278);
                cell?.SetCellValue(value);

                if (value != 0)
                {
                    _display1.Add(rowItem.Key, hitColumns);
                }
            }
        }

        public void Display2()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;

                var value = 0;
                if (_history.ContainsKey(rowIndex))
                {
                    for (var i = 0; i < 9; i++)
                    {
                        if (!_display1.ContainsKey(rowIndex) || !_display1[rowIndex].Select(u => u.Key).Contains(i))
                            continue;

                        var value1 = _history[rowIndex][i];
                        if (value1 != null)
                        {
                            var value2 = PreValue(GetPreRow(rowIndex), i);
                            if (value2?.Value == value1)
                                value++;
                        }
                    }
                }

                var cell = _sheet.GetRow(rowItem.Key)?.CreateCell(279);
                cell?.SetCellValue(value);
            }
        }

        public void Display3()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;

                var value = 0;
                if (_history.ContainsKey(rowIndex))
                {
                    for (var i = 0; i < 9; i++)
                    {
                        if (!_display1.ContainsKey(rowIndex) || !_display1[rowIndex].Select(u => u.Key).Contains(i))
                            continue;

                        var value1 = _history[rowIndex][i];
                        if (value1 != null)
                        {
                            var value2 = _display1[rowIndex].FirstOrDefault(u => u.Key == i).Value;
                            if (value2 != null && value2.Value != value1)
                                value++;
                        }
                    }
                }

                var cell = _sheet.GetRow(rowItem.Key)?.CreateCell(280);
                cell?.SetCellValue(value);
            }
        }

        private Dictionary<int, List<KeyValuePair<int, bool?>>> _display4 = new Dictionary<int, List<KeyValuePair<int, bool?>>>();
        public void Display4()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;

                var value = 0;
                var hitColumns = new List<KeyValuePair<int, bool?>>();
                rowIndex = GetPreRow(rowIndex);
                if (_history.ContainsKey(rowIndex))
                {
                    for (var i = 0; i < 9; i++)
                    {
                        var value1 = PreValue(rowIndex, i);
                        if (value1.HasValue)
                        {
                            var value2 = PreValue(GetPreRow(value1.Value.Key), i);
                            if (value2.HasValue && value1.Value.Value != value2?.Value)
                            {
                                hitColumns.Add(new KeyValuePair<int, bool?>(i, value1.Value.Value));
                                value++;
                            }

                        }
                    }
                }

                var cell = _sheet.GetRow(rowItem.Key)?.CreateCell(281);
                cell?.SetCellValue(value);

                if (value != 0)
                {
                    _display4.Add(rowItem.Key, hitColumns);
                }
            }
        }

        public void Display5()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;

                var value = 0;
                if (_history.ContainsKey(rowIndex))
                {
                    for (var i = 0; i < 9; i++)
                    {
                        if (!_display4.ContainsKey(rowIndex) || !_display4[rowIndex].Select(u => u.Key).Contains(i))
                            continue;

                        var value1 = _history[rowIndex][i];
                        if (value1 != null)
                        {
                            var value2 = PreValue(GetPreRow(rowIndex), i);
                            if (value2?.Value == value1)
                                value++;
                        }
                    }
                }

                var cell = _sheet.GetRow(rowItem.Key)?.CreateCell(282);
                cell?.SetCellValue(value);
            }
        }

        public void Display6()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;

                var value = 0;
                if (_history.ContainsKey(rowIndex))
                {
                    for (var i = 0; i < 9; i++)
                    {
                        if (!_display4.ContainsKey(rowIndex) || !_display4[rowIndex].Select(u => u.Key).Contains(i))
                            continue;

                        var value1 = _history[rowIndex][i];
                        if (value1 != null)
                        {
                            var value2 = _display4[rowIndex].FirstOrDefault(u => u.Key == i).Value;
                            if (value2 != null && value2.Value != value1)
                                value++;
                        }
                    }
                }

                var cell = _sheet.GetRow(rowItem.Key)?.CreateCell(283);
                cell?.SetCellValue(value);
            }
        }

        public void Display7()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;
                var row = _sheet.GetRow(rowIndex);
                if (row == null)
                {
                    continue;
                }
                var value1 = row?.GetCell(278)?.NumericCellValue??0;
                var value5 = row?.GetCell(281)?.NumericCellValue??0;
                var value3 = row?.GetCell(280)?.NumericCellValue??0;
                var cell = row?.CreateCell(284);
                var value = value1 + value5 - value3;
                cell?.SetCellValue(value);
            }
        }
        private Dictionary<int, List<KeyValuePair<int, bool?>>> _display7 = new Dictionary<int, List<KeyValuePair<int, bool?>>>();
        public void Display8()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;

                var value = 0;
                var hitColumns = new List<KeyValuePair<int, bool?>>();
                rowIndex = GetPreRow(rowIndex);
                if (_history.ContainsKey(rowIndex))
                {
                    for (var i = 0; i < 9; i++)
                    {
                        var value1 = PreValue(rowIndex, i);
                        if (value1.HasValue)
                        {
                            var value2 = PreValue(GetPreRow(value1.Value.Key), i);
                            if (value2 != null)
                            {
                                value2 = PreValue(GetPreRow(value2.Value.Key), i);
                                if (value1.Value.Value == value2?.Value)
                                {
                                    hitColumns.Add(new KeyValuePair<int, bool?>(i, value1.Value.Value));
                                    value++;
                                }
                            }
                        }
                    }
                }

                var cell = _sheet.GetRow(rowItem.Key)?.CreateCell(285);
                cell?.SetCellValue(value);

                if (value != 0)
                {
                    _display7.Add(rowItem.Key, hitColumns);
                }
            }
        }

        public void Display9()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;

                var value = 0;
                if (_history.ContainsKey(rowIndex))
                {
                    for (var i = 0; i < 9; i++)
                    {
                        if (!_display7.ContainsKey(rowIndex) || !_display7[rowIndex].Select(u => u.Key).Contains(i))
                            continue;

                        var value1 = _history[rowIndex][i];
                        if (value1 != null)
                        {
                            var value2 = PreValue(GetPreRow(rowIndex), i);
                            if (value2 != null)
                            {
                                value2 = PreValue(GetPreRow(value2.Value.Key), i);
                                if (value1.Value == value2?.Value)
                                {
                                    value++;
                                }
                            }

                        }
                    }
                }

                var cell = _sheet.GetRow(rowItem.Key)?.CreateCell(286);
                cell?.SetCellValue(value);
            }
        }

        public void Display10()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;

                var value = 0;
                if (_history.ContainsKey(rowIndex))
                {
                    for (var i = 0; i < 9; i++)
                    {
                        if (!_display7.ContainsKey(rowIndex) || !_display7[rowIndex].Select(u => u.Key).Contains(i))
                            continue;

                        var value1 = _history[rowIndex][i];
                        if (value1 != null)
                        {
                            var value2 = PreValue(GetPreRow(rowIndex), i);
                            if (value2 != null)
                            {
                                value2 = PreValue(GetPreRow(value2.Value.Key), i);
                                if (value1.Value != value2?.Value)
                                {
                                    value++;
                                }
                            }

                        }
                    }
                }

                var cell = _sheet.GetRow(rowItem.Key)?.CreateCell(287);
                cell?.SetCellValue(value);
            }
        }

        private Dictionary<int, List<KeyValuePair<int, bool?>>> _display10 = new Dictionary<int, List<KeyValuePair<int, bool?>>>();
        public void Display11()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;

                var value = 0;
                var hitColumns = new List<KeyValuePair<int, bool?>>();
                rowIndex = GetPreRow(rowIndex);
                if (_history.ContainsKey(rowIndex))
                {
                    for (var i = 0; i < 9; i++)
                    {
                        var value1 = PreValue(rowIndex, i);
                        if (value1.HasValue)
                        {
                            var value2 = PreValue(GetPreRow(value1.Value.Key), i);
                            if (value2 != null)
                            {
                                value2 = PreValue(GetPreRow(value2.Value.Key), i);
                                if (value1.Value.Value != value2?.Value)
                                {
                                    hitColumns.Add(new KeyValuePair<int, bool?>(i, value1.Value.Value));
                                    value++;
                                }
                            }
                        }
                    }
                }

                var cell = _sheet.GetRow(rowItem.Key)?.CreateCell(288);
                cell?.SetCellValue(value);

                if (value != 0)
                {
                    _display10.Add(rowItem.Key, hitColumns);
                }
            }
        }

        public void Display12()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;

                var value = 0;
                if (_history.ContainsKey(rowIndex))
                {
                    for (var i = 0; i < 9; i++)
                    {
                        if (!_display10.ContainsKey(rowIndex) || !_display10[rowIndex].Select(u => u.Key).Contains(i))
                            continue;

                        var value1 = _history[rowIndex][i];
                        if (value1 != null)
                        {
                            var value2 = PreValue(GetPreRow(rowIndex), i);
                            if (value2 != null)
                            {
                                value2 = PreValue(GetPreRow(value2.Value.Key), i);
                                if (value1.Value == value2?.Value)
                                {
                                    value++;
                                }
                            }

                        }
                    }
                }

                var cell = _sheet.GetRow(rowItem.Key)?.CreateCell(289);
                cell?.SetCellValue(value);
            }
           
        }

        public void Display13()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;

                var value = 0;
                if (_history.ContainsKey(rowIndex))
                {
                    for (var i = 0; i < 9; i++)
                    {
                        if (!_display10.ContainsKey(rowIndex) || !_display10[rowIndex].Select(u => u.Key).Contains(i))
                            continue;

                        var value1 = _history[rowIndex][i];
                        if (value1 != null)
                        {
                            var value2 = PreValue(GetPreRow(rowIndex), i);
                            if (value2 != null)
                            {
                                value2 = PreValue(GetPreRow(value2.Value.Key), i);
                                if (value1.Value != value2?.Value)
                                {
                                    value++;
                                }
                            }

                        }
                    }
                }

                var cell = _sheet.GetRow(rowItem.Key)?.CreateCell(290);
                cell?.SetCellValue(value);
            }
        }

        public void Display14()
        {
            foreach (var rowItem in _history)
            {

                var rowIndex = rowItem.Key;
                if (rowIndex > _lastRowIndex && _history[rowIndex].All(u => u == null))
                    continue;
                var row = _sheet.GetRow(rowIndex);
                if (row == null)
                {
                    continue;
                }
                var value1 = row.GetCell(285)?.NumericCellValue ?? 0;
                var value5 = row.GetCell(288)?.NumericCellValue ?? 0;
                var value3 = row.GetCell(287)?.NumericCellValue ?? 0;
                var cell = row?.CreateCell(291);
                var value = value1 + value5 - value3;
                cell?.SetCellValue(value);
            }
        }

        private int GetPreRow(int row)
        {
            if (row <= _lastRowIndex)
                return row - 1;

            return _lastRowIndex;
        }

        private KeyValuePair<int, bool>? PreValue(int row, int column)
        {
            if (row <= 1)
                return null;

            if (!_history.ContainsKey(row))
                return null;

            if (_history[row][column].HasValue)
                return new KeyValuePair<int, bool>(row, _history[row][column].Value);

            return PreValue(row - 1, column);
        }

        public void SetHeader()
        {

            var headerRow = _sheet.GetRow(0);
            for (int i = 0; i < 23; i++)
            {
                var cell = headerRow.CreateCell(i + 278);
                cell.SetCellValue(i + 1);
                _sheet.SetColumnWidth(i + 278, 700);
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
                    else if (cell.CellStyle.FillForegroundColor == 40)
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
