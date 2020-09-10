using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using HeProject.Model;
using HeProject.ProgressHandler.Common;
using NPOI.SS.UserModel;

namespace HeProject.ProgressHandler.CP
{
    public class ShaoGuHandler : ICPHandler
    {
        private ICellStyle shaoStyle;
        private ICellStyle guStyle;
        private int capacity;
        public CommonProcessContext Run(CommonProcessContext context)
        {
            var resultDic = context.ResultDic;
            capacity = resultDic[0].Capacity - 1;
            if (capacity < 1)
                return context;
            var sheet = context.Workbook.GetSheetAt(0);
            shaoStyle = context.Workbook.CreateCellStyle();
            shaoStyle.FillForegroundColor = 10;
            shaoStyle.FillPattern = FillPattern.SolidForeground;
            guStyle = context.Workbook.CreateCellStyle();
            guStyle.FillForegroundColor = 8;
            guStyle.FillPattern = FillPattern.SolidForeground;

            var lastSixResult = new Dictionary<int, int[]>();
            for (int i = 1; i <= 6; i++)
            {
                var rowValues = new int[6];
                var startIndex = 164;
                for (int j = startIndex; j < startIndex + rowValues.Length; j++)
                {
                    var cell = sheet.GetRow(capacity + i * 2).GetCell(j);
                    if (TryGetIntValue(cell, out int result))
                    {
                        rowValues[j - startIndex] = result;
                    }
                    else
                    {
                        rowValues[j - startIndex] = 0;
                    }
                }
                lastSixResult.Add(i - 1, rowValues);
            }
            ProcessShao(164, lastSixResult, context);
            ProcessGu(164, lastSixResult, context);
            lastSixResult = new Dictionary<int, int[]>();
            for (int i = 1; i <= 6; i++)
            {
                var rowValues = new int[6];
                var startIndex = 170;
                for (int j = startIndex; j < startIndex + rowValues.Length; j++)
                {
                    var cell = sheet.GetRow(capacity + i * 2).GetCell(j);
                    if (TryGetIntValue(cell, out int result))
                    {
                        rowValues[j - startIndex] = result;
                    }
                    else
                    {
                        rowValues[j - startIndex] = 0;
                    }
                }
                lastSixResult.Add(i - 1, rowValues);
            }
            ProcessShao(170, lastSixResult, context);
            ProcessGu(170, lastSixResult, context);
            lastSixResult = new Dictionary<int, int[]>();
            for (int i = 1; i <= 6; i++)
            {
                var rowValues = new int[3];
                var startIndex = 192;
                for (int j = startIndex; j < startIndex + rowValues.Length; j++)
                {
                    var cell = sheet.GetRow(capacity + i * 2).GetCell(j);
                    if (TryGetIntValue(cell, out int result))
                    {
                        rowValues[j - startIndex] = result;
                    }
                    else
                    {
                        rowValues[j - startIndex] = 0;
                    }
                }
                lastSixResult.Add(i - 1, rowValues);
            }
            ProcessShao(192, lastSixResult, context);
            ProcessGu(192, lastSixResult, context);
            lastSixResult = new Dictionary<int, int[]>();
            for (int i = 1; i <= 6; i++)
            {
                var rowValues = new int[3];
                var startIndex = 195;
                for (int j = startIndex; j < startIndex + rowValues.Length; j++)
                {
                    var cell = sheet.GetRow(capacity + i * 2).GetCell(j);
                    if (TryGetIntValue(cell, out int result))
                    {
                        rowValues[j - startIndex] = result;
                    }
                    else
                    {
                        rowValues[j - startIndex] = 0;
                    }
                }
                lastSixResult.Add(i - 1, rowValues);
            }
            ProcessShao(195, lastSixResult, context);
            ProcessGu(195, lastSixResult, context);
            lastSixResult = new Dictionary<int, int[]>();
            for (int i = 1; i <= 6; i++)
            {
                var rowValues = new int[3];
                var startIndex = 198;
                for (int j = startIndex; j < startIndex + rowValues.Length; j++)
                {
                    var cell = sheet.GetRow(capacity + i * 2).GetCell(j);
                    if (TryGetIntValue(cell, out int result))
                    {
                        rowValues[j - startIndex] = result;
                    }
                    else
                    {
                        rowValues[j - startIndex] = 0;
                    }
                }
                lastSixResult.Add(i - 1, rowValues);
            }
            ProcessShao(198, lastSixResult, context);
            ProcessGu(198, lastSixResult, context);
            lastSixResult = new Dictionary<int, int[]>();
            for (int i = 1; i <= 6; i++)
            {
                var rowValues = new int[6];
                var startIndex = 202;
                for (int j = startIndex; j < startIndex + rowValues.Length; j++)
                {
                    var cell = sheet.GetRow(capacity + i * 2).GetCell(j);
                    if (TryGetIntValue(cell, out int result))
                    {
                        rowValues[j - startIndex] = result;
                    }
                    else
                    {
                        rowValues[j - startIndex] = 0;
                    }
                }
                lastSixResult.Add(i - 1, rowValues);
            }
            ProcessShao(202, lastSixResult, context);
            lastSixResult = new Dictionary<int, int[]>();
            for (int i = 1; i <= 6; i++)
            {
                var rowValues = new int[1];
                var startIndex = 229;
                for (int j = startIndex; j < startIndex + rowValues.Length; j++)
                {
                    var cell = sheet.GetRow(capacity + i * 2).GetCell(j);
                    if (TryGetIntValue(cell, out int result))
                    {
                        rowValues[j - startIndex] = result;
                    }
                    else
                    {
                        rowValues[j - startIndex] = 0;
                    }
                }
                lastSixResult.Add(i - 1, rowValues);
            }
            ProcessGu(229, lastSixResult, context);
            lastSixResult = new Dictionary<int, int[]>();
            for (int i = 1; i <= 6; i++)
            {
                var rowValues = new int[1];
                var startIndex = 230;
                for (int j = startIndex; j < startIndex + rowValues.Length; j++)
                {
                    var cell = sheet.GetRow(capacity + i * 2).GetCell(j);
                    if (TryGetIntValue(cell, out int result))
                    {
                        rowValues[j - startIndex] = result;
                    }
                    else
                    {
                        rowValues[j - startIndex] = 0;
                    }
                }
                lastSixResult.Add(i - 1, rowValues);
            }
            ProcessGu(230, lastSixResult, context);
            lastSixResult = new Dictionary<int, int[]>();
            for (int i = 1; i <= 6; i++)
            {
                var rowValues = new int[1];
                var startIndex = 231;
                for (int j = startIndex; j < startIndex + rowValues.Length; j++)
                {
                    var cell = sheet.GetRow(capacity + i * 2).GetCell(j);
                    if (TryGetIntValue(cell, out int result))
                    {
                        rowValues[j - startIndex] = result;
                    }
                    else
                    {
                        rowValues[j - startIndex] = 0;
                    }
                }
                lastSixResult.Add(i - 1, rowValues);
            }
            ProcessGu(231, lastSixResult, context);
            var headerCell = sheet.GetRow(0).CreateCell(232);
            headerCell.SetCellValue("少");
            headerCell = sheet.GetRow(0).CreateCell(233);
            headerCell.SetCellValue("孤");
            sheet.SetColumnWidth(232, 1000);
            sheet.SetColumnWidth(233, 1000);
            for (int i = 1; i <= 6; i++)
            {
                var valueCell = sheet.GetRow(capacity + i * 2).CreateCell(232);
                valueCell.SetCellValue(context.Shao[i-1]);
                valueCell = sheet.GetRow(capacity + i * 2).CreateCell(233);
                valueCell.SetCellValue(context.Gu[i-1]);
            }
            return context;
        }

        private void ProcessShao(int startRowIndex, Dictionary<int, int[]> valueTable, CommonProcessContext context)
        {
            for (int i = 0; i < valueTable[0].Length; i++)
            {
                var values = new int[6];
                for (int j = 0; j < 6; j++)
                {
                    values[j] = valueTable[j][i];
                }
                ProcessShao(startRowIndex, i, values, context);
            }
        }

        private void ProcessShao(int startRowIndex, int columnIndex, int[] values, CommonProcessContext context)
        {
            var average = (double)values.Sum() / 6;
            var less = values.Count(u => ((double)u) < average);
            if (less == 2 || less == 1)
            {
                
                for (int i = 0; i < 6; i++)
                {
                    var cell = context.Workbook.GetSheetAt(0).GetRow(capacity + (i + 1) * 2).GetCell(startRowIndex + columnIndex);

                    if (TryGetIntValue(cell, out int value) && value < average)
                    {
                        context.Shao[i]++;
                        cell.CellStyle = shaoStyle;
                    }
                }
            }
            else if (less == 5 || less == 4)
            {
                for (int i = 0; i < 6; i++)
                {
                    var cell = context.Workbook.GetSheetAt(0).GetRow(capacity + (i + 1) * 2).GetCell(startRowIndex + columnIndex);

                    if (TryGetIntValue(cell, out int value) && value >= average)
                    {
                        context.Shao[i]++;
                        cell.CellStyle = shaoStyle;
                    }
                }
            }
        }

        private void ProcessGu(int startRowIndex, Dictionary<int, int[]> valueTable, CommonProcessContext context)
        {
            for (int i = 0; i < valueTable[0].Length; i++)
            {
                var values = new int[6];
                for (int j = 0; j < 6; j++)
                {
                    values[j] = valueTable[j][i];
                }
                ProcessGu(startRowIndex, i, values, context);
            }
        }

        private void ProcessGu(int startRowIndex, int columnIndex, int[] values, CommonProcessContext context)
        {
            var sum = values.Sum();
            if (sum >= 10)
            {
                var min = values.Min();
                if (values.Count(u => u == min) == 1)
                {
                    
                    var index = 0;
                    for (int i = 0; i < 6; i++)
                    {
                        if (values[i] == min)
                        {
                            index = i;
                            break;
                        }
                    }
                    context.Gu[index]++;
                    var sheet = context.Workbook.GetSheetAt(0);
                    var row = sheet.GetRow(capacity + (index + 1) * 2 + 1);
                    if (row == null)
                    {
                        row = sheet.CreateRow(capacity + (index + 1) * 2 + 1);
                        row.CreateCell(startRowIndex + columnIndex);
                    }
                    var cell = row.GetCell(startRowIndex + columnIndex);
                    if (cell == null)
                    {
                        cell = row.CreateCell(startRowIndex + columnIndex);
                    }

                    cell.CellStyle = guStyle;

                }
            }
            else
            {
                var max = values.Max();
                if (values.Count(u => u == max) == 1)
                {
                    var index = 0;
                    for (int i = 0; i < 6; i++)
                    {
                        if (values[i] == max)
                        {
                            index = i;
                            break;
                        }
                    }
                    context.Gu[index]++;
                    var sheet = context.Workbook.GetSheetAt(0);
                    var row = sheet.GetRow(capacity + (index + 1) * 2 + 1);
                    if (row == null)
                    {
                        row = sheet.CreateRow(capacity + (index + 1) * 2 + 1);
                        row.CreateCell(startRowIndex + columnIndex);
                    }
                    var cell = row.GetCell(startRowIndex + columnIndex);
                    if (cell == null)
                    {
                        cell = row.CreateCell(startRowIndex + columnIndex);
                    }
                    cell.CellStyle = guStyle;
                }
            }
        }

        private bool TryGetIntValue(ICell cell, out int value)
        {
            value = 0;
            if (cell.CellType == CellType.Numeric)
            {
                value = (int)cell.NumericCellValue;
                return true;
            }
            else if (cell.CellType == CellType.Blank)
            {
                return true;
            }
            return false;
        }

    }
}
