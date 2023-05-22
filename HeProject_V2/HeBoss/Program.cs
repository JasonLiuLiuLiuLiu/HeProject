using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using HeProject;
using HeProject._202305;
using HeProject.Model;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using static NPOI.HSSF.Util.HSSFColor;

namespace HeBoss
{
    internal class Program
    {
        static string tempDir = "temp";
        private static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("请勿关闭此窗口,正在处理中...");
                var seeds = SpliteFile("_Source.xlsx");
                if (seeds == null)
                {
                    return;
                }
                var maxFileCount = seeds.Length - 1;
                if (maxFileCount == 5)
                {
                    Console.WriteLine("输入文件至少5行");
                    return;
                }
                var width = maxFileCount.ToString().Length;
                var workbooks = new IWorkbook[maxFileCount + 1];
                for (int index = 0; index <= maxFileCount; index++)
                {
                    Console.WriteLine($"开始处理第{(index + 1).ToString().PadRight(width)}行数据,请不要删减或打开{tempDir}文件夹中的任何文件，以免影响程序运行....");
                    var resultDic = new Dictionary<int, ProcessContext>(11);
                    var sourcePath = tempDir + "/" + index + ".xlsx";

                    for (int i = 0; i < 6; i++)
                    {
                        var dataflow = new ProjectDataFlow(i);
                        var pipeline = dataflow.CreatePipeLine();
                        dataflow.Process(sourcePath);
                        pipeline.Wait();
                        resultDic.Add(i, dataflow.ProcessContext);
                    }
                    var writer = new WriteToExcel(resultDic);
                    var workbook = writer.ProcessDataflowResult();
                    if (workbook != null)
                    {
                        var commonContext = new CommonProcessContext(resultDic, workbook);
                        var commonFlow = new CommonProcessFlow();
                        commonFlow.Process(commonContext);
                        var resultPath = tempDir + "/out" + index + ".xlsx";
                        writer.SaveFile(workbook, resultPath);
                        workbooks[index] = workbook;
                    }
                }
                var maxWorkbook = workbooks[maxFileCount];
                var sheet = maxWorkbook.GetSheetAt(0);
                var daxiaotongji = sheet.GetRow(0).CreateCell(269);
                daxiaotongji.SetCellValue("大小统计");
                daxiaotongji.CellStyle = maxWorkbook.CreateCellStyle();
                daxiaotongji.CellStyle.Alignment = HorizontalAlignment.Center;
                sheet.AddMergedRegion(new CellRangeAddress(0, 0, 269, 269 + 8));

                for (int row = 2; row <= maxFileCount; row++)
                {
                    var rowIndex = (seeds[row] + 1) * 2 + row;
                    for (int index = 0; index < 9; index++)
                    {
                        var sourceCell = workbooks[row - 1].GetSheetAt(0).GetRow(rowIndex).GetCell(232 + index);
                        var targetCell = sheet.GetRow(row + 1).CreateCell(269 + index);
                        var style = maxWorkbook.CreateCellStyle();
                        style.FillForegroundColor = sourceCell.CellStyle.FillForegroundColor;
                        if (style.FillForegroundColor != 64)
                        {
                            style.FillPattern = FillPattern.SolidForeground;
                        }
                        targetCell.CellStyle = style;
                        targetCell.SetCellValue(sourceCell.NumericCellValue);
                    }
                }

                for (int row = 0; row < 6; row++)
                {
                    var rowIndex = (row + 1) * 2 + maxFileCount;
                    for (int index = 0; index < 9; index++)
                    {
                        var sourceCell = sheet.GetRow(rowIndex + 1).GetCell(232 + index);
                        var targetCell = sheet.GetRow(rowIndex + 1).CreateCell(269 + index);
                        targetCell.CellStyle = sourceCell.CellStyle;
                        targetCell.SetCellValue(sourceCell.NumericCellValue);
                    }
                }

                using (FileStream sw = File.Create("result.xlsx"))
                {
                    maxWorkbook.Write(sw);
                    sw.Close();
                }

                //XSSFWorkbook maxWorkbook;
                //using (FileStream file = new FileStream("result.xlsx", FileMode.Open, FileAccess.Read))
                //{
                //    maxWorkbook = new XSSFWorkbook(file);
                //}
                var p202305 = new P202305(maxWorkbook);
                p202305.Start();

                try
                {
                    var path = AppDomain.CurrentDomain.BaseDirectory + DateTime.Now.ToString("yyyy-MM-dd");
                    var fullPath = path + "\\" + DateTime.Now.ToString("HH-mm-ss-ffff") + ".xlsx";
                    if (!Directory.Exists(path))//如果不存在就创建file文件夹　　             　　
                        Directory.CreateDirectory(path);//创建该文件夹　　
                    using (FileStream sw = File.Create(fullPath))
                    {
                        maxWorkbook.Write(sw);
                        sw.Close();
                    }

                    Process.Start("explorer.exe", "/select, \"" + fullPath + "\"");
                    Process.Start(fullPath);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    Console.ReadKey();
                    throw;
                }

                if (Directory.Exists("temp"))
                {
                    Directory.Delete(tempDir, true);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private static int[] SpliteFile(string x)
        {
            try
            {
                if (x == null)
                {
                    Console.WriteLine("路径不允许为空!");
                    return null;
                }

                if (!File.Exists(x))
                {
                    Console.WriteLine("无法读取到输入文件:" + x + ",请检查文件是否存在!");
                    Console.ReadKey();
                    return null;
                }
                XSSFWorkbook hssfwb;
                using (FileStream file = new FileStream(x, FileMode.Open, FileAccess.Read))
                {
                    hssfwb = new XSSFWorkbook(file);
                }
                ISheet sheet = hssfwb.GetSheetAt(0);
                var seeds = new int[sheet.LastRowNum + 1];
                for (int row = 0; row <= sheet.LastRowNum; row++)
                {
                    if (sheet.GetRow(row) != null) //null is when the row only contains empty cells
                    {
                        if (!CheckSourceData(sheet.GetRow(row)))
                        {
                            Console.WriteLine($"检查到第{row}行数据格式有误,请关闭此程序并检查导入数据或清空表格重新导入!");
                            return null;
                        }

                        seeds[row] = sheet.GetRow(row).Where(u => !string.IsNullOrEmpty(u.ToString())).FirstOrDefault(u => ((int)u.NumericCellValue % 6) == u.ColumnIndex).ColumnIndex;

                    }
                }

                if (Directory.Exists("temp"))
                {
                    Directory.Delete(tempDir, true);
                }

                Directory.CreateDirectory(tempDir);

                IWorkbook[] workbooks = new IWorkbook[seeds.Length];
                ISheet[] sheets = new ISheet[seeds.Length];

                for (int row = 0; row < seeds.Length; row++)
                {
                    var workbook = new XSSFWorkbook();
                    sheets[row] = workbook.CreateSheet("Sheet1");
                    workbooks[row] = workbook;
                }

                for (int row = 0; row < seeds.Length; row++)
                {
                    var value = seeds[row];
                    for (int index = 0; index < sheets.Length; index++)
                    {
                        if (index < row)
                            continue;

                        var table = sheets[index];
                        var tableRow = table.CreateRow(row);
                        var cell = tableRow.CreateCell(value);
                        cell.SetCellValue(value);
                    }
                }

                for (int row = 0; row < seeds.Length; row++)
                {
                    var path = tempDir + "/" + row + ".xlsx";
                    using (FileStream sw = File.Create(path))
                    {
                        workbooks[row].Write(sw);
                        sw.Close();
                    }
                }

                return seeds;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }


        }

        private static bool CheckSourceData(IRow row)
        {
            var cells = row.Cells.Where(u => !string.IsNullOrEmpty(u.ToString())).Where(u => ((int)u.NumericCellValue % 6) == u.ColumnIndex);
            if (cells.Count() == 1)
                return true;
            return false;
        }
    }
}