using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks.Dataflow;
using HeProject.Model;

namespace HeProject.ProgressHandler.Source
{
    public class SourceHandler
    {
        private readonly ProcessContext _processContext;
        private readonly ExecutionDataflowBlockOptions _executionDataFlowBlockOptions;

        public SourceHandler(ProcessContext processContext)
        {
            _processContext = processContext;
            _executionDataFlowBlockOptions = new ExecutionDataflowBlockOptions()
            {
                MaxDegreeOfParallelism = Environment.ProcessorCount
            };
        }

        public IPropagatorBlock<int, int> CreatePaiXuBlock(int step)
        {
            var block = new TransformBlock<int, int>(row =>
            {
                List<int> order = new List<int>();
                bool[] handled = new bool[StepLength.SourceLength];
                if (row == 0)
                {
                    for (int i = 0; i < StepLength.SourceLength; i++)
                    {
                        if (handled[i])
                            continue;
                        order.Add(i);
                    }
                }
                else
                {
                    Handle(step, row, _processContext, order, handled);
                }

                for (int i = 0; i < StepLength.P3; i++)
                {
                    _processContext.SetSourceValue(step, row, order[i], i);
                }
                _processContext.SetSourceStepState(step, row, true);
                return row;
            }, _executionDataFlowBlockOptions);
            block.Completion.ContinueWith(t => { });
            return block;
        }

        public IPropagatorBlock<int, int> SetValueBlock(int step)
        {
            var block = new TransformBlock<int, int>(row =>
            {
                var beforeRow = _processContext.GetSourceRowResult(step - 1, row).FirstOrDefault(u => (bool)u.Value);
                var beforeRow_1 = _processContext.GetSourceRowResult(step - 1, row - 1).FirstOrDefault(u => (bool)u.Value);

                _processContext.SetSourceValue(step, row, (beforeRow_1.Key + beforeRow.Key) % 6, true);

                return row;
            }, _executionDataFlowBlockOptions);
            block.Completion.ContinueWith(t => { });
            return block;
        }

        public IPropagatorBlock<int, int> SetPart1ValueBlock()
        {
            var block = new TransformBlock<int, int>(row =>
            {
                var step3Value = _processContext.GetSourceRowResult(3, row).FirstOrDefault(u => (bool)u.Value);
                var step5Value = _processContext.GetSourceRowResult(5, row).FirstOrDefault(u => (bool)u.Value);
                var step7Value = _processContext.GetSourceRowResult(7, row).FirstOrDefault(u => (bool)u.Value);
                _processContext.SetP1Value(1, row, 0, step3Value.Key);
                _processContext.SetP1Value(1, row, 1, step5Value.Key);
                _processContext.SetP1Value(1, row, 2, step7Value.Key);
                return row;
            }, _executionDataFlowBlockOptions);
            block.Completion.ContinueWith(t => { });
            return block;
        }

        public IPropagatorBlock<int, int> SetPart2ValueBlock()
        {
            var block = new TransformBlock<int, int>(row =>
            {
                return row;
            }, _executionDataFlowBlockOptions);
            block.Completion.ContinueWith(t => { });
            return block;
        }

        public IPropagatorBlock<int, int> SetPart3ValueBlock()
        {
            var block = new TransformBlock<int, int>(row =>
            {
                _processContext.SetP1StepState(1, row, true);
                return row;
            }, _executionDataFlowBlockOptions);
            block.Completion.ContinueWith(t => { });
            return block;
        }

        public IPropagatorBlock<int, int> CreateDengJiBlock(int step)
        {
            var block = new TransformBlock<int, int>(row =>
            {
                try
                {
                    if (row < 1)
                    {
                        _processContext.SetSourceStepState(step, row, true);
                        return row;
                    }

                    var beforePaiXu = _processContext.GetP1RowResult(step - 2, row);
                    var paiXuResult = _processContext.GetP1RowResult(step - 1, row - 1);
                    for (int i = 0; i < StepLength.SourceLength; i++)
                    {
                        if ((bool)beforePaiXu[i])
                            _processContext.SetSourceValue(step, row, (int)paiXuResult[i], true);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }
                _processContext.SetSourceStepState(step, row, true);
                return row;
            }, _executionDataFlowBlockOptions);
            block.Completion.ContinueWith(t => { });
            return block;
        }

        private void Handle(int step, int row, ProcessContext context, List<int> order, bool[] handled, int[] columns = null)
        {
            var orderKeyValuePair = GetOrder(step, row, context, columns);
            if (orderKeyValuePair.Count == 0)
            {
                foreach (var column in columns)
                {
                    if (handled[column])
                        continue;
                    order.Add(column);
                    handled[column] = true;
                }
                return;
            }

            foreach (var valuePair in orderKeyValuePair)
            {
                if (handled[valuePair.Key])
                    continue;

                var sameCondition = orderKeyValuePair.Where(u => u.Value.Key == valuePair.Value.Key)
                    .Where(u => u.Value.Value == valuePair.Value.Value);
                if (sameCondition.Count() > 1)
                {
                    if (row < 1)
                        foreach (var column in columns)
                        {
                            if (handled[column])
                                continue;
                            order.Add(column);
                            handled[column] = true;
                        }
                    Handle(step, row - 1, context, order, handled, sameCondition.Select(u => u.Key).ToArray());
                }

                if (handled[valuePair.Key])
                    continue;
                order.Add(valuePair.Key);
                handled[valuePair.Key] = true;
            }
        }

        private List<KeyValuePair<int, KeyValuePair<int, int>>> GetOrder(int step, int row, ProcessContext context, int[] columns = null)
        {
            Dictionary<int, KeyValuePair<int, int>> distance = new Dictionary<int, KeyValuePair<int, int>>();
            for (int i = 0; i < StepLength.SourceLength; i++)
            {
                if (columns != null && !columns.Contains(i))
                    continue;

                int distanceLength = 0;
                int valueLength = 0;
                bool inValue = false;
                for (int j = row; j >= 0; j--)
                {
                    if (inValue)
                    {
                        if (context.GetSourceValue<bool>(step - 1, j, i))
                            valueLength++;
                        else
                        {
                            break;
                        }
                    }
                    else
                    {
                        if (context.GetSourceValue<bool>(step - 1, j, i))
                            inValue = true;
                        else
                        {
                            distanceLength++;
                        }
                    }

                    if (j == 0)
                        distanceLength = row;
                }
                distance.Add(i, new KeyValuePair<int, int>(distanceLength, valueLength));
            }
            return distance.OrderBy(u => u.Value.Key).ThenByDescending(u => u.Value.Value).ThenBy(u => u.Key).ToList();

            //for (int i = 0; i < StepLength.P3; i++)
            //{
            //    context.SetP1Value(3, row, i, afterOder.IndexOf(afterOder.FirstOrDefault(u => u.Key == i)));
            //}
            //return null;
        }
    }
}