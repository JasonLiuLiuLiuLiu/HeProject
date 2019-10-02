using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class P4HandleCommon
    {
        private int _step;
        public string GetOrder(int step, int row, ProcessContext context)
        {
            _step = step;
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
                Handle(row, context, order, handled);
            }

            var beforePaiXu = context.GetP4RowResult(step - 12, row);
            if (beforePaiXu == null || beforePaiXu.Count == 0)
                return null;

            var sum01 = 0;
            var sum23 = 0;
            var sum45 = 0;

            var indexs = beforePaiXu.Where(u => (int)u.Value>0);
            for (int i = 0; i <= StepLength.SourceLength; i++)
            {
                var value = indexs.Where(u => u.Key == order[i]).ToArray();
                if (value.Any())
                {
                    var valueCount = (int)value.FirstOrDefault().Value;
                    switch (order[i])
                    {
                        case 0:
                            sum01 += valueCount;
                            break;
                        case 1:
                            sum01 += valueCount;
                            break;
                        case 2:
                            sum23 += valueCount;
                            break;
                        case 3:
                            sum23 += valueCount;
                            break;
                        case 4:
                            sum45 += valueCount;
                            break;
                        case 5:
                            sum45 += valueCount;
                            break;
                    }
                }
            }
            context.SetP4Value(step,row,0,sum01);
            context.SetP4Value(step,row,1,sum23);
            context.SetP4Value(step,row,2,sum45);
            return null;
        }

        private void Handle(int row, ProcessContext context, List<int> order, bool[] handled, int[] columns = null)
        {
            var orderKeyValuePair = GetOrder(row, context, columns);
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

                var sameCondition = orderKeyValuePair.Where(u => u.Value.DistanceLength == valuePair.Value.DistanceLength)
                    .Where(u => u.Value.ValueLength == valuePair.Value.ValueLength).Where(u => u.Value.ValueLength == valuePair.Value.ValueCount).ToArray();
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
                    Handle(row - 1, context, order, handled, sameCondition.Select(u => u.Key).ToArray());
                }

                if (handled[valuePair.Key])
                    continue;
                order.Add(valuePair.Key);
                handled[valuePair.Key] = true;
            }
        }

        private List<KeyValuePair<int, OrderParams>> GetOrder(int row, ProcessContext context, int[] columns = null)
        {
            Dictionary<int, OrderParams> distance = new Dictionary<int, OrderParams>();
            for (int i = 0; i < StepLength.SourceLength; i++)
            {
                if (columns != null && !columns.Contains(i))
                    continue;

                int distanceLength = 0;
                int count = 0;
                int valueLength = 0;
                bool inValue = false;
                for (int j = row - 1; j >= 0; j--)
                {
                    count = context.GetP4Value<int>(_step - 12, j, i);

                    if (inValue)
                    {
                        if (count > 0)
                            valueLength++;
                        else
                        {
                            break;
                        }
                    }
                    else
                    {
                        if (count > 0)
                            inValue = true;
                        else
                        {
                            distanceLength++;
                        }
                    }

                    if (j == 0)
                        distanceLength = row - 1;

                }
                distance.Add(i, new OrderParams
                {
                    DistanceLength = distanceLength,
                    ValueLength = valueLength,
                    ValueCount = count
                });
            }
            return distance.OrderBy(u => u.Value.DistanceLength).ThenByDescending(u => u.Value.ValueLength).ThenByDescending(u => u.Value.ValueLength).ThenBy(u => u.Key).ToList();
        }

        public class OrderParams
        {
            public int DistanceLength { get; set; }
            public int ValueLength { get; set; }
            public int ValueCount { get; set; }
        }
    }
}
