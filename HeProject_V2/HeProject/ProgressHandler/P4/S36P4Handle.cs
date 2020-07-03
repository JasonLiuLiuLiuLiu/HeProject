using System;
using System.Collections.Generic;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S36P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            try
            {
                if (row < 6)
                {
                    return null;
                }

                var sumHistory = new Dictionary<int, object>[6, 7];
                for (int step = 1; step <= 6; step++)
                {
                    for (int rowIndex = 0; rowIndex <= 6; rowIndex++)
                    {
                        sumHistory[step - 1, rowIndex] = context.GetP4RowResult(step+18, row - rowIndex);
                    }
                }

                var mark = 0;
                for (int step = 0; step < 6; step++)
                {
                    for (int i = 0; i < 6; i++)
                    {
                        if (!sumHistory[step, 0].ContainsKey(i) || ((int)sumHistory[step, 0][i]) < 1)
                            continue;
                        for (int rowIndex = 1; rowIndex <= 6; rowIndex++)
                        {
                            if (sumHistory[step, rowIndex].ContainsKey(i) && ((int)sumHistory[step, rowIndex][i]) > 0)
                                break;
                            if (rowIndex == 5)
                            {
                                mark++;
                            }
                        }
                    }
                }
                context.SetP4Value(36, row, 0, mark);

                return null;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}
