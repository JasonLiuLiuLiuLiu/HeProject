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

        public SourceDataflow()
        {
            _executionDataFlowBlockOptions = new ExecutionDataflowBlockOptions()
            {
                MaxDegreeOfParallelism = Environment.ProcessorCount
            };
            _sourceHandle = new SourceHandler(ProcessContext);
        }

        public void Process(string filePath)
        {
            _startBlock.Post(filePath);
            _startBlock.Complete();
        }

        public Task CreatePipeline()
        {
            var step2PaiXuBlock = _sourceHandle.CreatePaiXuBlock(2);
            CreateReadFileBlock(step2PaiXuBlock);
            var step3DengJiBlock = _sourceHandle.CreateDengJiBlock(3);
            var step4PaiXuBlock = _sourceHandle.CreatePaiXuBlock(4);
            var step5DengJiBlock = _sourceHandle.CreateDengJiBlock(5);
            var step6PaiXuBlock = _sourceHandle.CreatePaiXuBlock(6);
            var step7DengJiBlock = _sourceHandle.CreateDengJiBlock(7);
            var setPart1Block = _sourceHandle.SetPart1ValueBlock();
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
                    for (int row = 0; row <= sheet.LastRowNum; row++)
                    {
                        if (sheet.GetRow(row) != null) //null is when the row only contains empty cells
                        {
                            if (!CheckSourceData(sheet.GetRow(row)))
                            {
                                Console.WriteLine($"检查到第{row}行数据格式有误,请关闭此程序并检查导入数据或清空表格重新导入!");
                                return;
                            }

                            for (int column = 0; column < StepLength.P1; column++)
                            {
                                var cellValue = sheet.GetRow(row).GetCell(column).ToString();
                                if (!string.IsNullOrEmpty(cellValue))
                                    ProcessContext.SetSourceValue(1, row, column, true);
                            }
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
            try
            {
                for (int j = 0; j < StepLength.P1; j++)
                {
                    var cellData = row.GetCell(j).ToString();
                    if (string.IsNullOrEmpty(cellData))
                        return false;
                    if (int.Parse(cellData) % 6 != j)
                        return false;
                }
            }
            catch
            {
                return false;
            }

            return true;
        }
    }
}