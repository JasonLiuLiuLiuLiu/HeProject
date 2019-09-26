using System;
using HeProject;

namespace HeBoss
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            //Console.WriteLine("当前软件为试用版,将会在2019年1月31日过期.");
            Console.WriteLine("请勿关闭此窗口,正在处理中...");
            // Thread.Sleep(40000);
            ////var dataflow = new ProjectDataFlow();
            ////var pipeline = dataflow.CreatePipeLine();
            ////dataflow.Process("_Input.xlsx");
            ////pipeline.Wait();
            ////WriteToExcel writer = new WriteToExcel(dataflow.ProcessContext);
            ////writer.Write();
            //Console.WriteLine("处理完成!");
            //Console.WriteLine($"输出结果保存在{Directory.GetCurrentDirectory()}\\_Output.xlsx");
            // Console.ReadKey();

            var sourceDataflow = new SourceDataflow();
            var pipeline = sourceDataflow.CreatePipeline();
            sourceDataflow.Process("_Source.xlsx");
            pipeline.Wait();
            // var result = JsonConvert.SerializeObject(sourceDataflow.ProcessContext);
            WriteToExcel writer = new WriteToExcel(sourceDataflow.ProcessContext);
            writer.Write();
        }
    }
}