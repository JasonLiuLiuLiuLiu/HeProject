using System.Collections.Generic;
using System.Linq;
using HeProject.Model;

namespace HeProject
{
    public class P3Handler : IProgressHandler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            Dictionary<int, KeyValuePair<int, int>> distance = new Dictionary<int, KeyValuePair<int, int>>();
            for (int i = 0; i < StepLength.P3; i++)
            {
                int distanceLength = 0;
                int valueLength = 0;
                bool inValue = false;
                for (int j = row; j >= 0; j--)
                {
                    if (inValue)
                    {
                        if (context.GetValue<bool>(2, j, i))
                            valueLength++;
                        else
                        {
                            break;
                        }
                    }
                    else
                    {
                        if (context.GetValue<bool>(2, j, i))
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
            var afterOder= distance.OrderBy(u => u.Value.Key).ThenByDescending(u => u.Value.Value).ThenBy(u=>u.Key).ToList();
           
            for (int i = 0; i < StepLength.P3; i++)
            {
                context.SetValue(3, row, i, afterOder.IndexOf(afterOder.FirstOrDefault(u=>u.Key==i)));
            }
            return null;
        }
    }
}
