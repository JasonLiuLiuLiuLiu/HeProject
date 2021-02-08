using System;
using System.Collections.Generic;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S47P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            try
            {
                if (row < 1)
                    return null;
                var p38Dic = context.GetP4RowResult(46, row - 1).CheckIntValue();
                var max3Num = 0;
                var max3Indexs = new List<int>();
                var max2Num = 0;
                var max2Indexs = new List<int>();

                for (int i = 0; i < 6; i++)
                {
                    if (((int)p38Dic[i]) >= 3)
                    {
                        max3Num++;
                        max3Indexs.Add(i);
                    }
                    if (((int)p38Dic[i]) >= 2)
                    {
                        max2Num++;
                        max2Indexs.Add(i);
                    }
                }
                var mark = 0;
                p38Dic = context.GetP4RowResult(46, row).CheckIntValue();
                if (max3Num == 2 && ((int)p38Dic[max3Indexs[0]] + (int)p38Dic[max3Indexs[1]]) >= 5)
                {
                    mark = 1;
                }
                if (max2Num == 3 && ((int)p38Dic[max2Indexs[0]] + (int)p38Dic[max2Indexs[1]] + (int)p38Dic[max2Indexs[2]]) >= 6)
                {
                    mark = 1;
                }
                if (max2Num == 4 && ((int)p38Dic[max2Indexs[0]] + (int)p38Dic[max2Indexs[1]] + (int)p38Dic[max2Indexs[2]] + (int)p38Dic[max2Indexs[3]]) >= 8)
                {
                    mark = 1;
                }
                context.SetP4Value(47, row, 0, mark);
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
