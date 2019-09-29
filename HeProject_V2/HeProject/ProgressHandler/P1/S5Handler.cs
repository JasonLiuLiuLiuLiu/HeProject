using HeProject.Model;
using System.Collections.Generic;
using System.Linq;
using HeProject.ProgressHandler.P1;

namespace HeProject
{
    public class S5Handler : IP1Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            List<int> order = new List<int>();
            bool[] handled = new bool[StepLength.P5];
            if (row == 0)
            {
                for (int i = 0; i < StepLength.P5; i++)
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

            for (int i = 0; i < StepLength.P3; i++)
            {
                context.SetP1Value(5, row, order[i], i);
            }
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
            Dictionary<int, KeyValuePair<int, int>> distance = new Dictionary<int, KeyValuePair<int, int>>();
            for (int i = 0; i < StepLength.P5; i++)
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
                        if (context.GetP1Value<bool>(4, j, i))
                            valueLength++;
                        else
                        {
                            break;
                        }
                    }
                    else
                    {
                        if (context.GetP1Value<bool>(4, j, i))
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
