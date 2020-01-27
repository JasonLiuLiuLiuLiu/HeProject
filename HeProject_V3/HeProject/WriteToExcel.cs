using System;
using System.Diagnostics;
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
        private ICellStyle _headerStyle;

        private ICellStyle _s0P1Style;
        private ICellStyle _s1P1Style;
        private ICellStyle _s2P1Style;


        public WriteToExcel(ProcessContext context)
        {
            _context = context;
        }

        public void Write()
        {
            try
            {
                IWorkbook workbook = new XSSFWorkbook();
                _headerStyle = workbook.CreateCellStyle();

                #region SetStyle

                _s0P1Style = workbook.CreateCellStyle();
                _s1P1Style = workbook.CreateCellStyle();
                _s2P1Style = workbook.CreateCellStyle();


                //https://www.cnblogs.com/love-zf/p/4874098.html
                _s0P1Style.FillForegroundColor = 19;
                _s1P1Style.FillForegroundColor = 13;
                _s2P1Style.FillForegroundColor = 22;


                _s0P1Style.FillPattern = FillPattern.SolidForeground;
                _s1P1Style.FillPattern = FillPattern.SolidForeground;
                _s2P1Style.FillPattern = FillPattern.SolidForeground;



                #endregion SetStyle

                ISheet sheet1 = workbook.CreateSheet("Sheet1");
                sheet1.DefaultColumnWidth = 1;
                sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
                SetHeader(sheet1, 0);
                for (int i = 1; i <= _context.Capacity; i++)
                {
                    IRow row = sheet1.CreateRow(i);
                    var index = i - 1;

                    #region P1

                    SetP1S0Value(row, index);
                    SetP1S1Value(row, index);
                    SetP1S2Value(row, index);

                    #endregion P1


                }
                SaveFile(workbook);
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

            var s0p1 = row.CreateCell(0);
            var s1p1 = row.CreateCell(StepLength.SourceLength * 1);
            var s2p1 = row.CreateCell(StepLength.SourceLength * 2);

            s0p1.SetCellValue("源数据");
            s1p1.SetCellValue("排序1");
            s2p1.SetCellValue("排序2");

            s0p1.CellStyle = _headerStyle;
            s1p1.CellStyle = _headerStyle;
            s2p1.CellStyle = _headerStyle;


            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, 0, StepLength.SourceLength * 1 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 1, StepLength.SourceLength * 2 - 1));
            sheet.AddMergedRegion(new CellRangeAddress(headerIndex, headerIndex, StepLength.SourceLength * 2, StepLength.SourceLength * 3 - 1));


        }

        #region SetP1Value

        private void SetP1S0Value(IRow row, int rowIndex)
        {
            int beforeColumn = 0;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP1Value<bool>(0, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s0P1Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        private void SetP1S1Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 1;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP1Value<bool>(1, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s1P1Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }



        private void SetP1S2Value(IRow row, int rowIndex)
        {
            int beforeColumn = StepLength.SourceLength * 2;
            for (int i = beforeColumn; i < StepLength.SourceLength + beforeColumn; i++)
            {
                var value = _context.GetP1Value<bool>(2, rowIndex, i - beforeColumn);
                var cell = row.CreateCell(i);
                cell.CellStyle = _s2P1Style;
                if (value)
                    cell.SetCellValue(i % StepLength.SourceLength);
            }
        }

        #endregion SetP1Value


        private void SaveFile(IWorkbook workBook)
        {
            try
            {
                var path = AppDomain.CurrentDomain.BaseDirectory + DateTime.Now.ToString("yyyy-MM-dd");
                var fullPath = path + "\\" + DateTime.Now.ToString("HH-mm-ss-ffff") + ".xlsx";
                if (!Directory.Exists(path))//如果不存在就创建file文件夹　　             　　
                    Directory.CreateDirectory(path);//创建该文件夹　　
                FileStream sw = File.Create(fullPath);
                workBook.Write(sw);
                sw.Close();
                // Process.Start("explorer.exe", "/select, \"" + fullPath + "\"");
                Process.Start(fullPath);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.ReadKey();
                throw;
            }
        }
    }
}