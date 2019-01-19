using System;
using System.IO;
using System.Threading;
using HeProject;

namespace HeBoss
{
    class Program
    {
        static void Main(string[] args)
        {
            var date = DateTime.Now;
            var expData = new DateTime(2019, 1, 31);
            if (date > expData)
            {
                Console.WriteLine("此软件已过期.");
                return;
            }
            Console.WriteLine("当前软件为试用版,将会在2019年1月31日过期.");
            Console.WriteLine("请勿关闭此窗口,正在处理中...");
            Thread.Sleep(40000);
            var dataflow = new ProjectDataFlow();
            var pipeline = dataflow.CreatePipeLine();
            dataflow.Process("_Input.xlsx");
            pipeline.Wait();
            WriteToExcel writer = new WriteToExcel(dataflow._processContext);
            writer.Write();
            Console.WriteLine("处理完成!");
            Console.WriteLine($"输出结果保存在{Directory.GetCurrentDirectory()}\\_Output.xlsx");
            Console.ReadKey();
        }
    }
}
