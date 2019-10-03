using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S25P4Handle : IP4Handler
    {
        private static int[] _ersan = new[] { 2, 3 };
        public string Hnalder(int row, ProcessContext context)
        {
            if (row == 0)
                return null;

            var currentValue = new[]
            {
                context.GetP4RowResult(7, row),
                context.GetP4RowResult(8, row),
                context.GetP4RowResult(9, row),
                context.GetP4RowResult(10, row),
                context.GetP4RowResult(11, row),
                context.GetP4RowResult(12, row),
                context.GetP4RowResult(13, row),
                context.GetP4RowResult(14, row),
                context.GetP4RowResult(15, row),
                context.GetP4RowResult(16, row),
                context.GetP4RowResult(17, row),
                context.GetP4RowResult(18, row)
                };

            var preValue = new[]
            {
                context.GetP4RowResult(7, row-1),
                context.GetP4RowResult(8, row-1),
                context.GetP4RowResult(9, row-1),
                context.GetP4RowResult(10, row-1),
                context.GetP4RowResult(11, row-1),
                context.GetP4RowResult(12, row-1),
                context.GetP4RowResult(13, row-1),
                context.GetP4RowResult(14, row-1),
                context.GetP4RowResult(15, row-1),
                context.GetP4RowResult(16, row-1),
                context.GetP4RowResult(17, row-1),
                context.GetP4RowResult(18, row-1)
                };

            int value = 0;

            for (int i = 0; i < 12; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    if (preValue[i].Count == 0 || !preValue[i].ContainsKey(j) || currentValue[i].Count == 0 ||
                        !currentValue[i].ContainsKey(j))
                    {
                        continue;
                    }

                    if ((int)preValue[i][j] == 0 && (int)currentValue[i][j] == (int)preValue[i][j])
                    {
                        value++;
                    }

                    if (_ersan.Contains((int)preValue[i][j]) && _ersan.Contains((int)currentValue[i][j]))
                    {
                        value++;
                    }
                }
            }

            context.SetP4Value(25, row, 0, value);

            if (row < 2)
                return null;

            var pre2Value = new[]
            {
                context.GetP4RowResult(7, row-2),
                context.GetP4RowResult(8, row-2),
                context.GetP4RowResult(9, row-2),
                context.GetP4RowResult(10, row-2),
                context.GetP4RowResult(11, row-2),
                context.GetP4RowResult(12, row-2),
                context.GetP4RowResult(13, row-2),
                context.GetP4RowResult(14, row-2),
                context.GetP4RowResult(15, row-2),
                context.GetP4RowResult(16, row-2),
                context.GetP4RowResult(17, row-2),
                context.GetP4RowResult(18, row-2)
            };

            int value2 = 0;

            for (int i = 0; i < 12; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    if (preValue[i].Count == 0 || !preValue[i].ContainsKey(j) || currentValue[i].Count == 0 ||
                        !currentValue[i].ContainsKey(j) || pre2Value[i].Count == 0 || !pre2Value[i].ContainsKey(j))
                    {
                        continue;
                    }

                    if ((int)preValue[i][j] == 0 && (int)currentValue[i][j] == (int)preValue[i][j] && (int)pre2Value[i][j] == (int)preValue[i][j])
                    {
                        value2++;
                    }

                    if (_ersan.Contains((int)preValue[i][j]) && _ersan.Contains((int)currentValue[i][j]) && _ersan.Contains((int)pre2Value[i][j]))
                    {
                        value2++;
                    }
                }
            }

            context.SetP4Value(25, row, 1, value2);

            return null;
        }
    }
}
