using System;
using HeProject;

namespace HeBoss
{
    class Program
    {
        static void Main(string[] args)
        {
            var dataflow = new ProjectDataFlow();
            var pipeline = dataflow.CreatePipeLine();
            dataflow.Process("input.xlsx");
            pipeline.Wait();
            Console.WriteLine("处理完成");
            Console.ReadKey();
        }
    }
}
