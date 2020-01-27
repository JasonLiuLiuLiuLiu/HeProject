using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using HeProject.Model;
using HeProject.ProgressHandler.P1;
using HeProject.ProgressHandler.P2;
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


        public async Task CreatePipeLine()
        {
            var sourceBroadCast = new BroadcastBlock<int>(i => i, _executionDataFlowBlockOptions);
            CreateStartBlock(sourceBroadCast);
            #region P1
            var p1CurrentBlock = CreateP1Block(1);
            sourceBroadCast.LinkTo(p1CurrentBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            for (var i = 2; i < 5; i++)
            {
                var newBlock = CreateP1Block(i);
                p1CurrentBlock.LinkTo(newBlock, new DataflowLinkOptions() { PropagateCompletion = true });
                p1CurrentBlock = newBlock;
            }
            #endregion

            #region P2

            var p2StartBlock = CreateP2Block(1);
            var p2CurrentBlock = p2StartBlock;
            for (int i = 2; i < 10; i++)
            {
                var newBlock = CreateP2Block(i);
                p2CurrentBlock.LinkTo(newBlock, new DataflowLinkOptions() { PropagateCompletion = true });
                p2CurrentBlock = newBlock;
            }
            #endregion

            var finallyP1Block = new ActionBlock<int>(x =>
            {
                var stepIndex = 7;
                for (int i = 0; i < 6; i++)
                {
                    if (ProcessContext.GetP1StepState(stepIndex + i, x))
                    {
                        p2StartBlock.Post(new P2BlockContext
                        {
                            Row = x,
                            Stage = i
                        });
                    }
                }

            });
            p1CurrentBlock.LinkTo(finallyP1Block, new DataflowLinkOptions() { PropagateCompletion = true });
            await finallyP1Block.Completion.ContinueWith(x =>
            {
                p2StartBlock.Complete();
            });

            var finallyP2Block = new ActionBlock<P2BlockContext>(x =>
              {
                //Console.WriteLine($"Stage:{x.Stage},Row:{x.Row}");
            });
            p2CurrentBlock.LinkTo(finallyP2Block, new DataflowLinkOptions() { PropagateCompletion = true });

            await finallyP2Block.Completion;
        }

        private void PrintState(ProgressState state)
        {
            //if (state.PNum == 4)
            if (state.Stage != 0)
                Console.WriteLine($"阶段{state.Stage}第{state.Step}步第{state.Row}行执行成功!");
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
                    var result = handler.Handler(x, ProcessContext);
                    //PrintState(new ProgressState(step, x) { ErrorMessage = result });
                    ProcessContext.SetP1StepState(step, x, true);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
                return x;
            }, _executionDataFlowBlockOptions);
            return progressBlock;
        }

        private IPropagatorBlock<P2BlockContext, P2BlockContext> CreateP2Block(int step)
        {
            var progressBlock = new TransformBlock<P2BlockContext, P2BlockContext>(x =>
            {
                try
                {
                    var handler = (IP2Handler)Activator.CreateInstance(Type.GetType($"HeProject.ProgressHandler.P2.P2S{step}Handler") ?? throw new InvalidOperationException());
                    var result = handler.Handler(x.Stage, x.Row, ProcessContext);
                    PrintState(new ProgressState(step, x.Row) { ErrorMessage = result, Stage = x.Stage });
                    ProcessContext.SetP2StepState(x.Stage, step, x.Row, true);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
                return x;
            }, _executionDataFlowBlockOptions);
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
                            if (!CheckSourceData(sheet.GetRow(row), out var message))
                            {
                                Console.WriteLine($"检查到第{row}行数据格式有误,{message},请关闭此程序并检查导入数据或清空表格重新导入!");
                                return;
                            }

                            var values = sheet.GetRow(row).Where(u => !string.IsNullOrEmpty(u.ToString()))
                                .Where(u => ((int)u.NumericCellValue % 11) == u.ColumnIndex).Select(u => u.ColumnIndex)
                                .ToArray();
                            foreach (var value in values)
                            {
                                ProcessContext.SetP1Value(0, row, value, true);
                            }
                            ProcessContext.SetP1StepState(0, row, true);
                            nextBlock.Post(row);
                        }
                        else
                        {
                            Console.WriteLine($"检查到第{row}行数据为空,请检查源数据!");
                            return;
                        }
                    }
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

        private bool CheckSourceData(IRow row, out string message)
        {
            var cells = row.Cells.Where(u => !string.IsNullOrEmpty(u.ToString())).Where(u => ((int)u.NumericCellValue % 11) == u.ColumnIndex).ToArray();
            if (!cells.Any())
            {
                message = "不存在任何数据,请检查源文件";
                return false;
            }

            if (cells.Count() == row.Cells.Count())
            {
                message = null;
                return true;
            }

            message = "存在错误数据";
            return false;
        }

        #endregion
    }
}