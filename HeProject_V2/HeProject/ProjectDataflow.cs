using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using HeProject.Model;
using HeProject.ProgressHandler.P1;
using HeProject.ProgressHandler.P2;
using HeProject.ProgressHandler.P3;
using HeProject.ProgressHandler.P4;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace HeProject
{
    public class ProjectDataFlow
    {
        private readonly ExecutionDataflowBlockOptions _executionDataFlowBlockOptions;
        public ProcessContext ProcessContext;
        private ITargetBlock<string> _startBlock;
        private int extendIndex = 0;

        public ProjectDataFlow(int index)
        {
            extendIndex = index;
            _executionDataFlowBlockOptions = new ExecutionDataflowBlockOptions()
            {
#if DEBUG
                MaxDegreeOfParallelism = 1
#else
                MaxDegreeOfParallelism = Environment.ProcessorCount
#endif

            };
        }
        public void Process(string filePath)
        {
            _startBlock.Post(filePath);
            _startBlock.Complete();
        }


        public async Task CreatePipeLine()
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

            #region P4

            var p4StartBlock = CreateP4Block(1);
            var currentP4Block = p4StartBlock;
            for (int i = 2; i <= 18; i++)
            {
                var block = CreateP4Block(i);
                currentP4Block.LinkTo(block, new DataflowLinkOptions() { PropagateCompletion = true });
                currentP4Block = block;
            }
            var s25P4Block = CreateP4Block(25);
            var s26P4Block = CreateP4Block(26);
            var s27P4Block = CreateP4Block(27);
            var s28P4Block = CreateP4Block(28);
            var s29P4Block = CreateP4Block(29);
            var s30P4Block = CreateP4Block(30);
            var s31P4Block = CreateP4Block(31);
            var s32P4Block = CreateP4Block(32);
            var s33P4Block = CreateP4Block(33);
            var s34P4Block = CreateP4Block(34);
            var s35P4Block = CreateP4Block(35);
            var s36P4Block = CreateP4Block(36);
            var s37P4Block = CreateP4Block(37);
            var s38P4Block = CreateP4Block(38);
            var s39P4Block = CreateP4Block(39);
            var s40P4Block = CreateP4Block(40);
            var s41P4Block = CreateP4Block(41);
            var s42P4Block = CreateP4Block(42);
            var s43P4Block = CreateP4Block(43);
            var s44P4Block = CreateP4Block(44);
            var s45P4Block = CreateP4Block(45);
            var s46P4Block = CreateP4Block(46);
            var s47P4Block = CreateP4Block(47);
            var s48P4Block = CreateP4Block(48);
            var s49P4Block = CreateP4Block(49);
            var s50P4Block = CreateP4Block(50);
            var s51P4Block = CreateP4Block(51);
            var s52P4Block = CreateP4Block(52);
            var s53P4Block = CreateP4Block(53);

            currentP4Block.LinkTo(s25P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s25P4Block.LinkTo(s26P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s26P4Block.LinkTo(s27P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s27P4Block.LinkTo(s28P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s28P4Block.LinkTo(s29P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s29P4Block.LinkTo(s30P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s30P4Block.LinkTo(s31P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s31P4Block.LinkTo(s32P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s32P4Block.LinkTo(s33P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s33P4Block.LinkTo(s34P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s34P4Block.LinkTo(s35P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s35P4Block.LinkTo(s36P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s36P4Block.LinkTo(s37P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s37P4Block.LinkTo(s38P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s38P4Block.LinkTo(s39P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s39P4Block.LinkTo(s40P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s40P4Block.LinkTo(s41P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s41P4Block.LinkTo(s42P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s42P4Block.LinkTo(s43P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s43P4Block.LinkTo(s44P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s44P4Block.LinkTo(s45P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s45P4Block.LinkTo(s46P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s46P4Block.LinkTo(s47P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s47P4Block.LinkTo(s48P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s48P4Block.LinkTo(s49P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s49P4Block.LinkTo(s50P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s50P4Block.LinkTo(s51P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s51P4Block.LinkTo(s52P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            s52P4Block.LinkTo(s53P4Block, new DataflowLinkOptions() { PropagateCompletion = true });
            var p4EndBlock = s53P4Block;
            var finallyP4Block = new ActionBlock<int>(x =>
            {
                // Console.WriteLine(x);
            });
            p4EndBlock.LinkTo(finallyP4Block, new DataflowLinkOptions() { PropagateCompletion = true });

            #endregion

            var sourceBroadCast = new BroadcastBlock<int>(i => i, _executionDataFlowBlockOptions);
            CreateStartBlock(sourceBroadCast);
            sourceBroadCast.LinkTo(s1P1Block, new DataflowLinkOptions() { PropagateCompletion = true });
            sourceBroadCast.LinkTo(s0P2Block, new DataflowLinkOptions() { PropagateCompletion = true });
            sourceBroadCast.LinkTo(s0P3Block, new DataflowLinkOptions() { PropagateCompletion = true });

            await Task.WhenAll(s3P1Block.Completion, s3P2Block.Completion, s3P3Block.Completion).ContinueWith(t =>
               {
                   for (int i = 0; i < ProcessContext.Capacity; i++)
                   {
                       p4StartBlock.Post(i);
                   }

                   p4StartBlock.Complete();
                   p4EndBlock.Completion.Wait();
               });
        }

        private void PrintState(ProgressState state)
        {
            //if (state.PNum == 4)
            //    Console.WriteLine($"第{state.Step}步执行成功!");
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

        private IPropagatorBlock<int, int> CreateP4Block(int step)
        {
            var progressBlock = new TransformBlock<int, int>(x =>
            {
                try
                {
                    var handler = (IP4Handler)Activator.CreateInstance(Type.GetType($"HeProject.ProgressHandler.P4.S{step}P4Handle") ?? throw new InvalidOperationException());
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
            progressBlock.Completion.ContinueWith(t => { PrintState(new ProgressState(step, -2, 4)); });
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
                    ProcessContext = new ProcessContext(sheet.LastRowNum + 2);
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

                    var lastRow = sheet.LastRowNum + 1;
                    ProcessContext.SetP1Value(0, lastRow, extendIndex, true);
                    ProcessContext.SetP1StepState(0, lastRow, true);
                    nextBlock.Post(lastRow);
                }
                catch (Exception)
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