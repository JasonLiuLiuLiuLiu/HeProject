using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using HeProject.Model;
using HeProject.ProgressHandler.Source;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace HeProject
{
    public class SourceDataflow
    {
        private readonly ExecutionDataflowBlockOptions _executionDataFlowBlockOptions;
        public ProcessContext ProcessContext;
        private ITargetBlock<string> _startBlock;
        private SourceHandler _sourceHandle;
        private ProjectDataFlow _projectDataFlow;

        public SourceDataflow()
        {
            _executionDataFlowBlockOptions = new ExecutionDataflowBlockOptions()
            {
                MaxDegreeOfParallelism = Environment.ProcessorCount
            };
            _sourceHandle = new SourceHandler();
            _projectDataFlow = new ProjectDataFlow();
        }

        public void Process(string filePath)
        {
            _startBlock.Post(filePath);
            _startBlock.Complete();
        }

        public Task CreatePipeline()
        {
            var pipeline = _projectDataFlow.CreatePipeLine();
            var step2PaiXuBlock = _sourceHandle.CreatePaiXuBlock(2);
            CreateReadFileBlock(step2PaiXuBlock);
            var step3DengJiBlock = _sourceHandle.CreateDengJiBlock(3);
            var step4PaiXuBlock = _sourceHandle.CreatePaiXuBlock(4);
            var step5DengJiBlock = _sourceHandle.CreateDengJiBlock(5);
            var step6PaiXuBlock = _sourceHandle.CreatePaiXuBlock(6);
            var step7DengJiBlock = _sourceHandle.CreateDengJiBlock(7);
            var setPart1Block = _sourceHandle.SetPart1ValueBlock();
            var step8SetValueBlock = _sourceHandle.SetValueBlock1(8);
            var step9PaiXuBlock = _sourceHandle.CreatePaiXuBlock(9);
            var step10DengJiBlock = _sourceHandle.CreateDengJiBlock(10);
            var step11PaiXuBlock = _sourceHandle.CreatePaiXuBlock(11);
            var step12DengJiBlock = _sourceHandle.CreateDengJiBlock(12);
            var step13PaiXuBlock = _sourceHandle.CreatePaiXuBlock(13);
            var step14DengJiBlock = _sourceHandle.CreateDengJiBlock(14);
            var setPart2Block = _sourceHandle.SetPart2ValueBlock();
            var step15SetValueBlock = _sourceHandle.SetValueBlock2(15);
            var step16PaiXuBlock = _sourceHandle.CreatePaiXuBlock(16);
            var step17DengJiBlock = _sourceHandle.CreateDengJiBlock(17);
            var step18PaiXuBlock = _sourceHandle.CreatePaiXuBlock(18);
            var step19DengJiBlock = _sourceHandle.CreateDengJiBlock(19);
            var step20PaiXuBlock = _sourceHandle.CreatePaiXuBlock(20);
            var step21DengJiBlock = _sourceHandle.CreateDengJiBlock(21);
            var setPart3Block = _sourceHandle.SetPart3ValueBlock();
            step2PaiXuBlock.LinkTo(step3DengJiBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step3DengJiBlock.LinkTo(step4PaiXuBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step4PaiXuBlock.LinkTo(step5DengJiBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step5DengJiBlock.LinkTo(step6PaiXuBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step6PaiXuBlock.LinkTo(step7DengJiBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step7DengJiBlock.LinkTo(setPart1Block, new DataflowLinkOptions() { PropagateCompletion = true });
            setPart1Block.LinkTo(step8SetValueBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step8SetValueBlock.LinkTo(step9PaiXuBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step9PaiXuBlock.LinkTo(step10DengJiBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step10DengJiBlock.LinkTo(step11PaiXuBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step11PaiXuBlock.LinkTo(step12DengJiBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step12DengJiBlock.LinkTo(step13PaiXuBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step13PaiXuBlock.LinkTo(step14DengJiBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step14DengJiBlock.LinkTo(setPart2Block, new DataflowLinkOptions() { PropagateCompletion = true });
            setPart2Block.LinkTo(step15SetValueBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step15SetValueBlock.LinkTo(step16PaiXuBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step16PaiXuBlock.LinkTo(step17DengJiBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step17DengJiBlock.LinkTo(step18PaiXuBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step18PaiXuBlock.LinkTo(step19DengJiBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step19DengJiBlock.LinkTo(step20PaiXuBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step20PaiXuBlock.LinkTo(step21DengJiBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            step21DengJiBlock.LinkTo(setPart3Block, new DataflowLinkOptions() { PropagateCompletion = true });
            var finallyBlock = new ActionBlock<int>(x => { _projectDataFlow.StartBlock.Post(x); });
            finallyBlock.Completion.ContinueWith(x => { _projectDataFlow.StartBlock.Complete(); });
            setPart3Block.LinkTo(finallyBlock, new DataflowLinkOptions() { PropagateCompletion = true });
            return pipeline;
        }

        private void CreateReadFileBlock(IPropagatorBlock<int, int> s2P1Block)
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
                    _sourceHandle.ProcessContext = ProcessContext;
                    _projectDataFlow.ProcessContext = ProcessContext;
                    for (int row = 0; row <= sheet.LastRowNum; row++)
                    {
                        if (sheet.GetRow(row) != null) //null is when the row only contains empty cells
                        {
                            if (!CheckSourceData(sheet.GetRow(row)))
                            {
                                Console.WriteLine($"检查到第{row}行数据格式有误,请关闭此程序并检查导入数据或清空表格重新导入!");
                                return;
                            }

                            ProcessContext.SetSourceValue(1, row, sheet.GetRow(row).Where(u => !string.IsNullOrEmpty(u.ToString())).FirstOrDefault(u => ((int)u.NumericCellValue % 6) == u.ColumnIndex).ColumnIndex, true);
                            ProcessContext.SetSourceStepState(1, row, true);
                            s2P1Block.Post(row);
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
                s2P1Block.Complete();
            });
        }

        private bool CheckSourceData(IRow row)
        {
            var cells = row.Cells.Where(u => !string.IsNullOrEmpty(u.ToString())).Where(u => ((int)u.NumericCellValue % 6) == u.ColumnIndex);
            if (cells.Count() == 1)
                return true;
            return false;
        }
    }
}