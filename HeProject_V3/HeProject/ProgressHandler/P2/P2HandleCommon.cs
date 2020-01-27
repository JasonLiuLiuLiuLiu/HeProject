using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2HandleCommon
    {
        private int _step;
        private int _stage;
        private int _offset;
        public string GetOrder(int stage, int step, int row, ProcessContext p2Context,int offset)
        {
            try
            {
                _step = step;
                _stage = stage;
                _offset = offset;
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
                    Handle(row, p2Context, order, handled);
                }

                var beforePaiXu = p2Context.GetP2RowResult(_stage, step - offset, row);
                if (beforePaiXu == null || beforePaiXu.Count == 0)
                    return null;

                var indexs = beforePaiXu.Where(u => (bool)u.Value).Select(u => u.Key).ToArray();
                for (int i = 0; i < StepLength.SourceLength; i++)
                {
                    if (indexs.Contains(order[i]))
                    {
                        p2Context.SetP2Value(_stage, step, row, i, true);
                    }
                }

                return null;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        private void Handle(int row, ProcessContext p2Context, List<int> order, bool[] handled, int[] columns = null)
        {
            var orderKeyValuePair = GetOrder(row, p2Context, columns);
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
                    Handle(row - 1, p2Context, order, handled, sameCondition.Select(u => u.Key).ToArray());
                }

                if (handled[valuePair.Key])
                    continue;
                order.Add(valuePair.Key);
                handled[valuePair.Key] = true;
            }
        }

        private List<KeyValuePair<int, KeyValuePair<int, int>>> GetOrder(int row, ProcessContext p2Context, int[] columns = null)
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
                            if (p2Context.GetP2Value<bool>(_stage, _step - _offset, j, i))
                                valueLength++;
                            else
                            {
                                break;
                            }
                        }
                        else
                        {
                            if (p2Context.GetP2Value<bool>(_stage, _step - _offset, j, i))
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
