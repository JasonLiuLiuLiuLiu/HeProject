using HeProject.Model;
using HeProject.ProgressHandler.Common;
using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace HeProject.ProgressHandler.CP
{
    public class BigSmallChangeColor : ICPHandler
    {
        public CommonProcessContext Run(CommonProcessContext context)
        {
            var sheet = context.Workbook.GetSheetAt(0);
            var resultDic = context.ResultDic;
            var capacity = resultDic[0].Capacity - 1;
            var row = sheet.GetRow(0);
            var red = context.Workbook.CreateCellStyle();
            red.FillForegroundColor = 10;
            red.FillPattern = FillPattern.SolidForeground;
            var blue = context.Workbook.CreateCellStyle();
            blue.FillForegroundColor = 40;
            blue.FillPattern = FillPattern.SolidForeground;
            var lastSixResult = new Dictionary<int, int[]>();
            var startIndex = 232;
            for (int i = 1; i <= 6; i++)
            {
                var rowValues = new int[9];
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

            for (int i = 0; i < 9; i++)
            {
                var maxValue = lastSixResult[0][i];
                var minValue = lastSixResult[0][i];
                var maxIndex = new List<int>();
                var minIndex = new List<int>();
                for (int j = 0; j < 6; j++)
                {
                    var value = lastSixResult[j][i];
                    if (value > maxValue)
                    {
                        maxIndex.Clear();
                        maxValue = value;
                        maxIndex.Add(j);
                    }
                    else if (value == maxValue)
                    {
                        maxIndex.Add(j);
                    }

                    if (value < minValue)
                    {
                        minIndex.Clear();
                        minValue = value;
                        minIndex.Add(j);
                    }
                    else if (value == minValue)
                    {
                        minIndex.Add(j);
                    }
                }

                for (int j = 0; j < 6; j++)
                {
                    var cell = sheet.GetRow(capacity + (j + 1) * 2).GetCell(i + startIndex);
                    if (maxIndex.Count == 6 || minIndex.Count == 6)
                    {
                        cell.CellStyle = null;
                    }
                    else if (maxIndex.Contains(j))
                    {
                        cell.CellStyle = red;
                    }
                    else if (minIndex.Contains(j))
                    {
                        cell.CellStyle = blue;
                    }
                    else
                    {
                        cell.CellStyle = null;
                    }
                }
            }

            return context;
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
