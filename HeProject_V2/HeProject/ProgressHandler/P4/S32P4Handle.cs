using System;
using System.Collections.Generic;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S32P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            if (row < 4)
            {
                return null;
            }

            try
            {
                var resultDic = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
                var result = new int[3];

                for (int i = 1; i < 7; i++)
                {
                    resultDic.Add(i, new Dictionary<int, Dictionary<int, object>>());
                    for (int j = 0; j < 4; j++)
                    {
                        resultDic[i][j] = new Dictionary<int, object>();
                        resultDic[i][j] = context.GetP4RowResult(i, row - j);
                    }
                }

                for (int i = 1; i < 7; i++)
                {
                    var indexs = resultDic[i][0].Where(u => ((int)u.Value) != 0);
                    foreach (var item in indexs)
                    {
                        var index = item.Key;
                        if (resultDic[i][1].ContainsKey(index) && (int)resultDic[i][1][index] != 0)
                        {
                            result[0] += (int)item.Value;
                        }
                        else if ((resultDic[i][2].ContainsKey(index) && (int)resultDic[i][2][index] != 0) || (resultDic[i][3].ContainsKey(index) && (int)resultDic[i][3][index] != 0))
                        {
                            result[1] += (int)item.Value;
                        }
                        else
                        {
                            result[2] += (int)item.Value;
                        }
                    }
                }

                context.SetP4Value(32, row, 0, result[0]);
                context.SetP4Value(32, row, 1, result[1]);
                context.SetP4Value(32, row, 2, result[2]);

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
