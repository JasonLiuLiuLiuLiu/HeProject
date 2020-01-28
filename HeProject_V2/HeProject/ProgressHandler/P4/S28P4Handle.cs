using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S28P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            if (row <= 3)
            {
                for (int i = 0; i < 4; i++)
                {
                    context.SetP4Value(28, row, i, 0);
                }

                return null;
            }

            var result = new int[4];

            for (int i = 19; i < 25; i++)
            {
                var valueDic0 = context.GetP4RowResult(i, row);
                var valueDic1 = context.GetP4RowResult(i, row - 1);
                var valueDic2 = context.GetP4RowResult(i, row - 2);
                var valueDic3 = context.GetP4RowResult(i, row - 3);

                for (int j = 0; j < 6; j++)
                {
                    if (!valueDic0.ContainsKey(j)) continue;

                    if (!valueDic1.ContainsKey(j))
                    {
                        result[0]++;
                        continue;
                    }

                    if (!valueDic2.ContainsKey(j))
                    {
                        result[1]++;
                        continue;
                    }

                    if (!valueDic3.ContainsKey(j))
                    {
                        result[2]++;
                        continue;
                    }

                    result[3]++;
                }
            }

            for (int i = 0; i < 4; i++)
            {
                context.SetP4Value(28, row, i, result[i]);
            }

            return null;
        }
    }
}
