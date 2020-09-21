using HeProject.Model;
using HeProject.ProgressHandler.Common;
using NPOI.SS.UserModel;

namespace HeProject.ProgressHandler.CP
{
    public class CopyHandler : ICPHandler
    {
        public CommonProcessContext Run(CommonProcessContext context)
        {
            var sheet = context.Workbook.GetSheetAt(0);
            var row = sheet.GetRow(0);
            var sumStyle = context.Workbook.CreateCellStyle();
            sumStyle.FillForegroundColor = 43;
            sumStyle.FillPattern = FillPattern.SolidForeground;
            var sumStyle2 = context.Workbook.CreateCellStyle();
            sumStyle2.FillForegroundColor = 22;
            sumStyle2.FillPattern = FillPattern.SolidForeground;
            var cell = row.CreateCell(234);
            cell.SetCellValue("II");
            cell = row.CreateCell(235);
            cell.SetCellValue("跳1");
            cell = row.CreateCell(239);
            cell.SetCellValue("黄1");
            sheet.SetColumnWidth(234, 700);
            sheet.SetColumnWidth(235, 700);
            sheet.SetColumnWidth(239, 700);
            for (int i = 1; i < context.ResultDic[0].Capacity; i++)
            {
                row = sheet.GetRow(i);
                cell = row.CreateCell(234);
                cell.SetCellValue(row.GetCell(162).NumericCellValue);
                cell.CellStyle = sumStyle;
                cell = row.CreateCell(235);
                cell.SetCellValue(row.GetCell(176).NumericCellValue);
                cell.CellStyle = sumStyle2;
                cell = row.CreateCell(239);
                cell.SetCellValue(row.GetCell(208).NumericCellValue);
                cell.CellStyle = sumStyle2;
            }
            for (int i = context.ResultDic[0].Capacity+1; i < context.ResultDic[0].Capacity+12; i+=2)
            {
                row = sheet.GetRow(i);
                cell = row.CreateCell(234);
                cell.SetCellValue(row.GetCell(162).NumericCellValue);
                cell.CellStyle = sumStyle;
                cell = row.CreateCell(235);
                cell.SetCellValue(row.GetCell(176).NumericCellValue);
                cell.CellStyle = sumStyle2;
                cell = row.CreateCell(239);
                cell.SetCellValue(row.GetCell(208).NumericCellValue);
                cell.CellStyle = sumStyle2;
            }
            return context;
        }
    }
}
