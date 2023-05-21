using System;
using System.Collections.Generic;
using System.Linq;


namespace HeProject._202305
{
    internal class HandleCommon
    {
        private const int SourceLength = 9;
        private Func<int, int> _getPreRow;

        public HandleCommon(Func<int, int> getPreRow)
        {
            _getPreRow = getPreRow;
        }
        public List<int> GetOrder(int row, Dictionary<int, bool?[]> history, bool? big)
        {
            List<int> order = new List<int>();
            bool[] handled = new bool[9];
            if (row == 1)
            {
                for (int i = 0; i < SourceLength; i++)
                {
                    if (handled[i])
                        continue;
                    order.Add(i);
                }
            }
            else
            {
                Handle(row, history, order, handled, big);
            }

            return order;
        }

        private void Handle(int row, Dictionary<int, bool?[]> history, List<int> order, bool[] handled, bool? big, int[] columns = null)
        {
            var orderKeyValuePair = GetOrder(row, history, big, columns);
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
                    Handle(_getPreRow(row), history, order, handled, big, sameCondition.Select(u => u.Key).ToArray());
                }

                if (handled[valuePair.Key])
                    continue;
                order.Add(valuePair.Key);
                handled[valuePair.Key] = true;
            }
        }

        private List<KeyValuePair<int, KeyValuePair<int, int>>> GetOrder(int row, Dictionary<int, bool?[]> history, bool? big, int[] columns = null)
        {
            try
            {
                Dictionary<int, KeyValuePair<int, int>> distance = new Dictionary<int, KeyValuePair<int, int>>();
                for (int i = 0; i < SourceLength; i++)
                {
                    if (columns != null && !columns.Contains(i))
                        continue;

                    int distanceLength = 0;
                    int valueLength = 0;
                    bool inValue = false;
                    for (int j = _getPreRow(row); j >= 1; j--)
                    {
                        if (inValue)
                        {
                            if (history.ContainsKey(j) && (big.HasValue ? (history[j][i].HasValue && history[j][i].Value == big) : history[j][i].HasValue))
                                valueLength++;
                            else
                            {
                                break;
                            }
                        }
                        else
                        {
                            if (history.ContainsKey(j) && (big.HasValue ? (history[j][i].HasValue && history[j][i].Value == big) : history[j][i].HasValue))
                                inValue = true;
                            else
                            {
                                distanceLength++;
                            }
                        }

                        if (j == 1)
                            distanceLength = row - 1;

                    }

                    distance.Add(i, new KeyValuePair<int, int>(distanceLength, valueLength));
                }

                return distance.OrderBy(u => u.Value.Key).ThenByDescending(u => u.Value.Value).ThenBy(u => u.Key)
                    .ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
