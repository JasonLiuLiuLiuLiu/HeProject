using System;
using System.IO;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using HeProject.Model;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace HeProject
{
    public class ProjectDataFlow
    {
        private readonly ExecutionDataflowBlockOptions _executionDataflowBlockOptions;
        public ProcessContext _processContext;
        private int _totalRow;
        private ITargetBlock<string> _startBlock;
        private object _lock = new object();

        public ProjectDataFlow()
        {

            _executionDataflowBlockOptions = new ExecutionDataflowBlockOptions()
            {
                MaxDegreeOfParallelism = Environment.ProcessorCount
            };
        }

        public void Process(string filePath)
        {
            _startBlock.Post(filePath);
            _startBlock.Complete();
        }

        public Task CreatePipeLine()
        {
            var currentBlock = CreateReadFileBlock(CreateProcessBlock(2));
            for (int i = 3; i < 10; i++)
            {
                var block = CreateProcessBlock(i);
                currentBlock.LinkTo(block, new DataflowLinkOptions() { PropagateCompletion = true });
                currentBlock = block;
            }
            var finallyBlock = new ActionBlock<int>(x => Console.WriteLine($"第{x}行处理完成!"));
            currentBlock.LinkTo(finallyBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            return finallyBlock.Completion;
        }

        private void PrintState(ProgressState state)
        {
            Task.Run(() =>
            {
                lock (_lock)
                {
                    Console.SetCursorPosition(0, state.Step);
                    if (state.Row == -1)
                    {
                        Console.WriteLine($"第{state.Step}步执行失败:{state.ErrorMessage}!");
                    }
                    //else if (state.Row == -2)
                    //{
                    //    Console.WriteLine($"第{state.Step}步执行成功!");
                    //}
                }
            });
        }

        private IPropagatorBlock<int, int> CreateReadFileBlock(IPropagatorBlock<int, int> p2Block)
        {

            _startBlock = new ActionBlock<string>(x =>
                {
                    try
                    {
                        if (x == null)
                        {
                            PrintState(new ProgressState(1, -1) { ErrorMessage = "路径不允许为空!" });
                            return;
                        }

                        if (!File.Exists(x))
                        {
                            PrintState(new ProgressState(1, -1) { ErrorMessage = "文件不存在!" });
                            return;
                        }
                        XSSFWorkbook hssfwb;
                        using (FileStream file = new FileStream(x, FileMode.Open, FileAccess.Read))
                        {
                            hssfwb = new XSSFWorkbook(file);
                        }
                        ISheet sheet = hssfwb.GetSheetAt(0);
                        _totalRow = sheet.LastRowNum;
                        _processContext = new ProcessContext(sheet.LastRowNum + 1);
                        for (int row = 0; row <= sheet.LastRowNum; row++)
                        {
                            if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                            {
                                for (int column = 0; column < StepLength.P1; column++)
                                {
                                    _processContext.SetValue(1, row, column, (int)sheet.GetRow(row).GetCell(column).NumericCellValue);
                                }
                                _processContext.SetStepState(1, row, true);
                                p2Block.Post(row);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }



                }, _executionDataflowBlockOptions);
            _startBlock.Completion.ContinueWith(x =>
            {
                PrintState(new ProgressState(1, -2));
                p2Block.Complete();
            });
            return p2Block;
        }

        private IPropagatorBlock<int, int> CreateProcessBlock(int step)
        {
            var progressBlock = new TransformBlock<int, int>(x =>
            {
                try
                {
                    var handler = (IProgressHandler)Activator.CreateInstance(Type.GetType($"HeProject.P{step}Handler") ?? throw new InvalidOperationException());
                    var result = handler.Hnalder(x, _processContext);
                    if (result != null)
                        PrintState(new ProgressState(step, -1) { ErrorMessage = result });
                    _processContext.SetStepState(step, x, true);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
                return x;
            }, _executionDataflowBlockOptions);
            progressBlock.Completion.ContinueWith(t => { PrintState(new ProgressState(step, -2)); });
            return progressBlock;
        }
    }
}
