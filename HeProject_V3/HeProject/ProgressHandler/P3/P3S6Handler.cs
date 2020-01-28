using System;
using HeProject.Model;

namespace HeProject.ProgressHandler.P3
{
    public class P3S6Handler : IP3Handler
    {
        public string Handler(int row, ProcessContext context)
        {
            var result = new int[4];
            for (var i = 2; i < 28; i += 3)
            {
                var values = context.GetP2RowResult(5, i, row);
                for (int j = 0; j < 4; j++)
                {
                    if ((bool)values[j])
                    {
                        result[j]++;
                    }
                }
            }
            for (int i = 0; i < 4; i++)
            {
                context.SetP3Value(6, row, i, result[i]);
            }
            return null;
        }
    }
}
