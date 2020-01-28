using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using HeProject;
using HeProject.Model;
using Newtonsoft.Json;

namespace HeBoss
{
    internal class Program
    {
        private static void Main()
        {
            try
            {
                Console.WriteLine("请勿关闭此窗口,正在处理中...");
                var resultDic = new Dictionary<int, ProcessContext>(11);
                Stopwatch sp = new Stopwatch();
                sp.Start();
                for (int i = 0; i < 11; i++)
                {
                    var dataflow = new ProjectDataFlow(i);
                    var pipeline = dataflow.CreatePipeLine();
                    dataflow.Process("_Source.xlsx");
                    pipeline.Wait();
                    resultDic.Add(i, dataflow.ProcessContext);
                }
                sp.Stop();
                Console.WriteLine(sp.ElapsedMilliseconds);

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