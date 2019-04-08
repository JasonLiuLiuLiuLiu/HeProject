using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using HeProject.Model;
using HeProject.Part2;
using HeProject.ProgressHandler.P3;
using HeProject.ProgressHandler.P4;
using HeProject.ProgressHandler.P5;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace HeProject
{
    public class ProjectDataFlow
    {
        private readonly ExecutionDataflowBlockOptions _executionDataFlowBlockOptions;
        public ProcessContext ProcessContext;
        private int _totalRow;
        public BroadcastBlock<int> StartBlock;

        public ProjectDataFlow()
        {
            _executionDataFlowBlockOptions = new ExecutionDataflowBlockOptions()
            {
                MaxDegreeOfParallelism = Environment.ProcessorCount
            };
        }

        //public void Process(string filePath)
        //{
        //    _startBlock.Post(filePath);
        //    _startBlock.Complete();
        //}

        public Task CreatePipeLine()
        {
            #region Start

            var s2P1Block = CreateP1Block(2);
            var s2P3Block = CreateP3Block(2);
            StartBlock = new BroadcastBlock<int>(i =>
            {
                //Console.WriteLine($"第一,二部分第{i}行开始处理;");
                return i;
            }, _executionDataFlowBlockOptions);
            StartBlock.LinkTo(s2P1Block, new DataflowLinkOptions() { PropagateCompletion = true });
            StartBlock.LinkTo(s2P3Block, new DataflowLinkOptions() { PropagateCompletion = true });

            #endregion Start

            //  CreateReadFileBlock(s2P1Block, s2P3Block);

            #region P1 P2

            var p1S3Block = CreateP1Block(3);
            var currentP1Block = p1S3Block;
            for (int i = 4; i < 11; i++)
            {
                var block = CreateP1Block(i);
                currentP1Block.LinkTo(block, new DataflowLinkOptions() { PropagateCompletion = true });
                currentP1Block = block;
            }

            IPropagatorBlock<int, int> p2S2Block = CreateP2Block(2);
            var currentP2Block = p2S2Block;
            for (int i = 3; i < 11; i++)
            {
                var block = CreateP2Block(i);
                currentP2Block.LinkTo(block, new DataflowLinkOptions() { PropagateCompletion = true });
                currentP2Block = block;
            }

            var s2P1BroadcastBlock = new BroadcastBlock<int>(i =>
            {
                //Console.WriteLine($"第一,二部分第{i}行开始处理;");
                return i;
            }, _executionDataFlowBlockOptions);
            s2P1Block.LinkTo(s2P1BroadcastBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            s2P1BroadcastBlock.LinkTo(p1S3Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s2P1BroadcastBlock.LinkTo(p2S2Block, new DataflowLinkOptions() { PropagateCompletion = true });

            var finallyP1Block = new ActionBlock<int>(x =>
            {
                // Console.WriteLine($"第一部分{x}行处理完成!");
                // Thread.Sleep(2000);
            });
            currentP1Block.LinkTo(finallyP1Block, new DataflowLinkOptions() { PropagateCompletion = true });

            var finallyP2Block = new ActionBlock<int>(x =>
            {
                // Console.WriteLine($"第二部分{x}行处理完成!");
                //  Thread.Sleep(2000);
            });
            currentP2Block.LinkTo(finallyP2Block, new DataflowLinkOptions() { PropagateCompletion = true });

            #endregion P1 P2

            #region P3 P4

            var p3S3Block = CreateP3Block(3);
            var currentP3Block = p3S3Block;
            for (int i = 4; i < 11; i++)
            {
                var block = CreateP3Block(i);
                currentP3Block.LinkTo(block, new DataflowLinkOptions() { PropagateCompletion = true });
                currentP3Block = block;
            }

            IPropagatorBlock<int, int> p4S2Block = CreateP4Block(2);
            var currentP4Block = p4S2Block;
            for (int i = 3; i < 11; i++)
            {
                var block = CreateP4Block(i);
                currentP4Block.LinkTo(block, new DataflowLinkOptions() { PropagateCompletion = true });
                currentP4Block = block;
            }

            var s2P3BroadcastBlock = new BroadcastBlock<int>(i =>
            {
                //Console.WriteLine($"第一,二部分第{i}行开始处理;");
                return i;
            }, _executionDataFlowBlockOptions);
            s2P3Block.LinkTo(s2P3BroadcastBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            s2P3BroadcastBlock.LinkTo(p3S3Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s2P3BroadcastBlock.LinkTo(p4S2Block, new DataflowLinkOptions() { PropagateCompletion = true });

            var finallyP3Block = new ActionBlock<int>(x =>
            {
                //  Console.WriteLine($"第三部分{x}行处理完成!");
                // Thread.Sleep(2000);
            });
            currentP3Block.LinkTo(finallyP3Block, new DataflowLinkOptions() { PropagateCompletion = true });

            var finallyP4Block = new ActionBlock<int>(x =>
            {
                // Console.WriteLine($"第四部分{x}行处理完成!");
                //  Thread.Sleep(2000);
            });
            currentP4Block.LinkTo(finallyP4Block, new DataflowLinkOptions() { PropagateCompletion = true });

            #endregion P3 P4

            return Task.WhenAll(new Task[] { finallyP1Block.Completion, finallyP2Block.Completion, finallyP3Block.Completion, finallyP4Block.Completion }).ContinueWith(t =>
            {
                for (int i = 0; i < ProcessContext.Capacity; i++)
                {
                    new S1P5Handler().Hnalder(i, ProcessContext);
                    new S2P5Handler().Hnalder(i, ProcessContext)
                }
            });
        }

        private void PrintState(ProgressState state)
        {
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

        //private void CreateReadFileBlock(IPropagatorBlock<int, int> s2P1Block, IPropagatorBlock<int, int> s2P3Block)
        //{
        //    _startBlock = new ActionBlock<string>(x =>
        //        {
        //            try
        //            {
        //                if (x == null)
        //                {
        //                    PrintState(new ProgressState(1, -1) { ErrorMessage = "路径不允许为空!" });
        //                    return;
        //                }

        //                if (!File.Exists(x))
        //                {
        //                    Console.WriteLine("无法读取到输入文件:" + x + ",请检查文件是否存在!");
        //                    Console.ReadKey();
        //                    return;
        //                }
        //                XSSFWorkbook hssfwb;
        //                using (FileStream file = new FileStream(x, FileMode.Open, FileAccess.Read))
        //                {
        //                    hssfwb = new XSSFWorkbook(file);
        //                }
        //                ISheet sheet = hssfwb.GetSheetAt(0);
        //                _totalRow = sheet.LastRowNum;
        //                ProcessContext = new ProcessContext(sheet.LastRowNum + 1);
        //                for (int row = 0; row <= sheet.LastRowNum; row++)
        //                {
        //                    if (sheet.GetRow(row) != null) //null is when the row only contains empty cells
        //                    {
        //                        if (!CheckSourceData(sheet.GetRow(row)))
        //                        {
        //                            Console.WriteLine($"检查到第{row}行数据格式有误,请关闭此程序并检查导入数据或清空表格重新导入!");
        //                            return;
        //                        }

        //                        for (int column = 0; column < StepLength.P1; column++)
        //                        {
        //                            ProcessContext.SetP1Value(1, row, column, (int)sheet.GetRow(row).GetCell(column).NumericCellValue);
        //                        }
        //                        ProcessContext.SetP1StepState(1, row, true);
        //                        s2P1Block.Post(row);
        //                        s2P3Block.Post(row);
        //                    }
        //                }
        //            }
        //            catch (Exception e)
        //            {
        //                Console.WriteLine("输入文件被占用,请关闭该文件!");
        //                Console.ReadKey();
        //            }
        //        }, _executionDataFlowBlockOptions);
        //    _startBlock.Completion.ContinueWith(x =>
        //    {
        //        PrintState(new ProgressState(1, -2));
        //        s2P1Block.Complete();
        //        s2P3Block.Complete();
        //    });
        //}

        private bool CheckSourceData(IRow row)
        {
            try
            {
                int countOfZero = 0;
                for (int j = 0; j < StepLength.P1; j++)
                {
                    var cellData = row.GetCell(j).ToString();
                    if (string.IsNullOrEmpty(cellData))
                        return false;
                    if (int.Parse(cellData) == 0)
                        countOfZero++;
                    if (countOfZero > 6)
                        return false;
                }
            }
            catch
            {
                return false;
            }

            return true;
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
            progressBlock.Completion.ContinueWith(t => { PrintState(new ProgressState(step, -2)); });
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

        private IPropagatorBlock<int, int> CreateP4Block(int step)
        {
            var progressBlock = new TransformBlock<int, int>(x =>
            {
                try
                {
                    var handler = (IP4Handler)Activator.CreateInstance(Type.GetType($"HeProject.ProgressHandler.P4.S{step}P4Handler") ?? throw new InvalidOperationException());
                    var result = handler.Hnalder(x, ProcessContext);
                    if (result != null)
                        PrintState(new ProgressState(step, -1) { ErrorMessage = result });
                    ProcessContext.SetP4StepState(step, x, true);
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
    }
}