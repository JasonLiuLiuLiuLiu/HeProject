using System;
using System.Collections.Generic;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P1
{
    public class P1HandleCommon
    {
        private int _step;
        public string GetOrder(int step, int row, ProcessContext context)
        {
            try
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

                var beforePaiXu = context.GetP1RowResult(step - 1, row);
                if (beforePaiXu == null || beforePaiXu.Count == 0)
                    return null;

                var indexs = beforePaiXu.Where(u => (bool)u.Value).Select(u => u.Key);
                for (int i = 0; i < StepLength.SourceLength; i++)
                {
                    if (indexs.Contains(order[i]))
                    {
                        context.SetP1Value(step, row, i, true);
                    }
                }

                for (int i = 0; i < StepLength.SourceLength; i++)
                {
                    context.SetP1Value(step + 3, row, order[i], i);
                }
                context.SetP1StepState(step + 3, row, true);

                return null;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
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

                var sameCondition = orderKeyValuePair.Where(u => u.Value.Key == valuePair.Value.Key)
                    .Where(u => u.Value.Value == valuePair.Value.Value).ToArray();
                if (sameCondition.Length > 1)
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

        private List<KeyValuePair<int, KeyValuePair<int, int>>> GetOrder(int row, ProcessContext context, int[] columns = null)
        {
            try
            {
                Dictionary<int, KeyValuePair<int, int>> distance = new Dictionary<int, KeyValuePair<int, int>>();
                for (int i = 0; i < StepLength.SourceLength; i++)
                {
                    if (columns != null && !columns.Contains(i))
                        continue;

                    int distanceLength = 0;
                    int valueLength = 0;
                    bool inValue = false;
                    for (int j = row - 1; j >= 0; j--)
                    {
                        if (inValue)
                        {
                            if (context.GetP1Value<bool>(_step - 1, j, i))
                                valueLength++;
                            else
                            {
                                break;
                            }
                        }
                        else
                        {
                            if (context.GetP1Value<bool>(_step - 1, j, i))
                                inValue = true;
                            else
                            {
                                distanceLength++;
                            }
                        }

                        if (j == 0)
                            distanceLength = row - 1;

                    }

                    distance.Add(i, new KeyValuePair<int, int>(distanceLength, valueLength));
                }

                return distance.OrderBy(u => u.Value.Key).ThenByDescending(u => u.Value.Value).ThenBy(u => u.Key)
                    .ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }
    }
}
