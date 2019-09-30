using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using HeProject.Model;
using HeProject.ProgressHandler.P1;
using HeProject.ProgressHandler.P2;
using HeProject.ProgressHandler.P3;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace HeProject
{
    public class ProjectDataFlow
    {
        private readonly ExecutionDataflowBlockOptions _executionDataFlowBlockOptions;
        public ProcessContext ProcessContext;
        private ITargetBlock<string> _startBlock;

        public ProjectDataFlow()
        {
            _executionDataFlowBlockOptions = new ExecutionDataflowBlockOptions()
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
            #region P1

            var s1P1Block = CreateP1Block(1);
            var s2P1Block = CreateP1Block(2);
            var s3P1Block = CreateP1Block(3);
            s1P1Block.LinkTo(s2P1Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s2P1Block.LinkTo(s3P1Block, new DataflowLinkOptions() { PropagateCompletion = true });
            var finallyP1Block = new ActionBlock<int>(x =>
            {
               // Console.WriteLine(x);
            });
            s3P1Block.LinkTo(finallyP1Block, new DataflowLinkOptions() { PropagateCompletion = true });

            #endregion

            #region P2

            var s0P2Block = CreateP2Block(0);
            var s1P2Block = CreateP2Block(1);
            var s2P2Block = CreateP2Block(2);
            var s3P2Block = CreateP2Block(3);
            s0P2Block.LinkTo(s1P2Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s1P2Block.LinkTo(s2P2Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s2P2Block.LinkTo(s3P2Block, new DataflowLinkOptions() { PropagateCompletion = true });
            var finallyP2Block = new ActionBlock<int>(x =>
            {
                // Console.WriteLine(x);
            });
            s3P2Block.LinkTo(finallyP2Block, new DataflowLinkOptions() { PropagateCompletion = true });

            #endregion

            #region P3

            var s0P3Block = CreateP3Block(0);
            var s1P3Block = CreateP3Block(1);
            var s2P3Block = CreateP3Block(2);
            var s3P3Block = CreateP3Block(3);
            s0P3Block.LinkTo(s1P3Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s1P3Block.LinkTo(s2P3Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s2P3Block.LinkTo(s3P3Block, new DataflowLinkOptions() { PropagateCompletion = true });
            var finallyP3Block = new ActionBlock<int>(x =>
            {
                // Console.WriteLine(x);
            });
            s3P3Block.LinkTo(finallyP3Block, new DataflowLinkOptions() { PropagateCompletion = true });

            #endregion
            var sourceBroadCast = new BroadcastBlock<int>(i => i, _executionDataFlowBlockOptions);
            CreateStartBlock(sourceBroadCast);
            sourceBroadCast.LinkTo(s1P1Block, new DataflowLinkOptions() { PropagateCompletion = true });
            sourceBroadCast.LinkTo(s0P2Block, new DataflowLinkOptions() { PropagateCompletion = true });
            sourceBroadCast.LinkTo(s0P3Block, new DataflowLinkOptions() { PropagateCompletion = true });

            return Task.WhenAll(new[] {s3P1Block.Completion, s3P2Block.Completion, s3P3Block.Completion});
        }

        private void PrintState(ProgressState state)
        {
            Console.WriteLine($"第{state.Step}步执行成功!");
            //Task.Run(() =>
            //{
            //    lock (_lock)
            //    {
            //        Console.SetCursorPosition(0, state.Step);
            //        if (state.Row == -1)
            //        {
            //            Console.WriteLine($"第{state.Step}步执行失败:{state.ErrorMessage}!");
            //        }
            //        //else if (state.Row == -2)
            //        //{
            //        //    Console.WriteLine($"第{state.Step}步执行成功!");
            //        //}
            //    }
            //});
        }




        #region CreateBlock

        private IPropagatorBlock<int, int> CreateP1Block(int step)
        {
            var progressBlock = new TransformBlock<int, int>(x =>
            {
                try
                {
                    var handler = (IP1Handler)Activator.CreateInstance(Type.GetType($"HeProject.S{step}Handler") ?? throw new InvalidOperationException());
                    var result = handler.Hnalder(x, ProcessContext);
                    if (result != null)
                        PrintState(new ProgressState(step, -1) { ErrorMessage = result });
                    ProcessContext.SetP1StepState(step, x, true);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
                return x;
            }, _executionDataFlowBlockOptions);
            progressBlock.Completion.ContinueWith(t => { PrintState(new ProgressState(step, int.MinValue)); });
            return progressBlock;
        }

        private IPropagatorBlock<int, int> CreateP2Block(int step)
        {
            var progressBlock = new TransformBlock<int, int>(x =>
            {
                try
                {
                    var handler = (IP2Handler)Activator.CreateInstance(Type.GetType($"HeProject.ProgressHandler.P2.S{step}P2Handler") ?? throw new InvalidOperationException());
                    var result = handler.Hnalder(x, ProcessContext);
                    if (result != null)
                        PrintState(new ProgressState(step, -1) { ErrorMessage = result });
                    ProcessContext.SetP2StepState(step, x, true);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
                return x;
            }, _executionDataFlowBlockOptions);
            progressBlock.Completion.ContinueWith(t => { PrintState(new ProgressState(step, -2)); });
            return progressBlock;
        }

        private IPropagatorBlock<int, int> CreateP3Block(int step)
        {
            var progressBlock = new TransformBlock<int, int>(x =>
            {
                try
                {
                    var handler = (IP3Handler)Activator.CreateInstance(Type.GetType($"HeProject.ProgressHandler.P3.S{step}P3Handler") ?? throw new InvalidOperationException());
                    var result = handler.Hnalder(x, ProcessContext);
                    if (result != null)
                        PrintState(new ProgressState(step, -1) { ErrorMessage = result });
                    ProcessContext.SetP3StepState(step, x, true);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
                return x;
            }, _executionDataFlowBlockOptions);
            progressBlock.Completion.ContinueWith(t => { PrintState(new ProgressState(step, -2)); });
            return progressBlock;
        }

        #endregion CreateBlock

        #region ReadExcel

        private void CreateStartBlock(IPropagatorBlock<int, int> nextBlock)
        {
            _startBlock = new ActionBlock<string>(x =>
            {
                try
                {
                    if (x == null)
                    {
                        Console.WriteLine("路径不允许为空!");
                        return;
                    }

                    if (!File.Exists(x))
                    {
                        Console.WriteLine("无法读取到输入文件:" + x + ",请检查文件是否存在!");
                        Console.ReadKey();
                        return;
                    }
                    XSSFWorkbook hssfwb;
                    using (FileStream file = new FileStream(x, FileMode.Open, FileAccess.Read))
                    {
                        hssfwb = new XSSFWorkbook(file);
                    }
                    ISheet sheet = hssfwb.GetSheetAt(0);
                    ProcessContext = new ProcessContext(sheet.LastRowNum + 1);
                    for (int row = 0; row <= sheet.LastRowNum; row++)
                    {
                        if (sheet.GetRow(row) != null) //null is when the row only contains empty cells
                        {
                            if (!CheckSourceData(sheet.GetRow(row)))
                            {
                                Console.WriteLine($"检查到第{row}行数据格式有误,请关闭此程序并检查导入数据或清空表格重新导入!");
                                return;
                            }

                            ProcessContext.SetP1Value(0, row, sheet.GetRow(row).Where(u => !string.IsNullOrEmpty(u.ToString())).FirstOrDefault(u => ((int)u.NumericCellValue % 6) == u.ColumnIndex).ColumnIndex, true);
                            ProcessContext.SetP1StepState(0, row, true);
                            nextBlock.Post(row);
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("输入文件被占用,请关闭该文件!");
                    Console.ReadKey();
                }
            }, _executionDataFlowBlockOptions);
            _startBlock.Completion.ContinueWith(x =>
            {
                nextBlock.Complete();
            });
        }

        private bool CheckSourceData(IRow row)
        {
            var cells = row.Cells.Where(u => !string.IsNullOrEmpty(u.ToString())).Where(u => ((int)u.NumericCellValue % 6) == u.ColumnIndex);
            if (cells.Count() == 1)
                return true;
            return false;
        }

        #endregion
    }
}