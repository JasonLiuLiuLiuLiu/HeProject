using System.Collections.Generic;
using HeProject.Model;
using HeProject.ProgressHandler.Common;
using NPOI.SS.UserModel;

namespace HeProject.ProgressHandler.CP
{
    public class TenBingAndSmall : ICPHandler
    {
        public CommonProcessContext Run(CommonProcessContext context)
        {
            var sheet = context.Workbook.GetSheetAt(0);
            var lastSix = new int[6][];
            var index = 0;
            var row = sheet.GetRow(0);
            var sumStyle = context.Workbook.CreateCellStyle();
            sumStyle.FillForegroundColor = 13;
            sumStyle.FillPattern = FillPattern.SolidForeground;
            var sumStyle2 = context.Workbook.CreateCellStyle();
            sumStyle2.FillForegroundColor = 45;
            sumStyle2.FillPattern = FillPattern.SolidForeground;
            var sumStyle3 = context.Workbook.CreateCellStyle();
            sumStyle3.FillForegroundColor = 43;
            sumStyle3.FillPattern = FillPattern.SolidForeground;
            var cell = row.CreateCell(236);
            cell.SetCellValue("10大");
            cell = row.CreateCell(237);
            cell.SetCellValue("10小");
            cell = row.CreateCell(238);
            cell.SetCellValue("热0");
            sheet.SetColumnWidth(236, 700);
            sheet.SetColumnWidth(237, 700);
            sheet.SetColumnWidth(238, 700);
            for (int i = context.ResultDic[0].Capacity + 1; i < context.ResultDic[0].Capacity + 12; i += 2)
            {
                row = sheet.GetRow(i);
                lastSix[index] = new int[7];
                lastSix[index][0] = (int)row.GetCell(178).NumericCellValue;
                lastSix[index][1] = (int)row.GetCell(179).NumericCellValue;
                lastSix[index][2] = (int)row.GetCell(198).NumericCellValue;
                lastSix[index][3] = (int)row.GetCell(199).NumericCellValue;
                lastSix[index][4] = (int)row.GetCell(200).NumericCellValue;
                lastSix[index][5] = (int)row.GetCell(229).NumericCellValue;
                lastSix[index][6] = (int)row.GetCell(230).NumericCellValue;
                index++;
            }
            List<int> bingIndex = new List<int>(7);
            for (int i = 0; i < 7; i++)
            {
                var sum = 0;
                for (int j = 0; j < 6; j++)
                {
                    sum += lastSix[j][i];
                }

                if (sum > 9)
                {
                    bingIndex.Add(i);
                }
            }

            var tenBing = new int[6];
            var tenSmall = new int[6];
            for (int i = 0; i < 7; i++)
            {
                if (bingIndex.Contains(i))
                {
                    for (int j = 0; j < 6; j++)
                    {
                        tenBing[j] += lastSix[j][i];
                    }
                }
                else
                {
                    for (int j = 0; j < 6; j++)
                    {
                        tenSmall[j] += lastSix[j][i];
                    }
                }

            }

            index = 0;
            for (int i = context.ResultDic[0].Capacity + 1; i < context.ResultDic[0].Capacity + 12; i += 2)
            {
                row = sheet.GetRow(i);
                cell = row.CreateCell(236);
                cell.SetCellValue(tenBing[index]);
                cell.CellStyle = sumStyle;
                cell = row.CreateCell(237);
                cell.SetCellValue(tenSmall[index]);
                cell.CellStyle = sumStyle2;
                cell = row.CreateCell(238);
                cell.SetCellValue(row.GetCell(194).NumericCellValue + row.GetCell(197).NumericCellValue);
                cell.CellStyle = sumStyle3;
                index++;
            }
            return context;
        }
    }
}
