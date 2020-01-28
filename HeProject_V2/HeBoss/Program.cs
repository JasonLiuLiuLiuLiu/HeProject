using System;
using System.Collections.Generic;
using System.Diagnostics;
using HeProject;
using HeProject.Model;

namespace HeBoss
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("请勿关闭此窗口,正在处理中...");
                var resultDic = new Dictionary<int, ProcessContext>(11);

                for (int i = 0; i < 6; i++)
                {
                    var dataflow = new ProjectDataFlow(i);
                    var pipeline = dataflow.CreatePipeLine();
                    dataflow.Process("_Source.xlsx");
                    pipeline.Wait();
                    resultDic.Add(i, dataflow.ProcessContext);
                }
                var writer = new WriteToExcel(resultDic);
                writer.Write();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }
    }
}