using System;
using System.Collections.Generic;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S37P4Handle : IP4Handler
    {
        private static int[] P4S32Yellow = new[] { 6, 7, 8 };
        public string Hnalder(int row, ProcessContext context)
        {
            try
            {
                if (row < 1)
                    return null;
                var p4S32ValueDic = row > 4 ? context.GetP4RowResult(32, row) : new Dictionary<int, object>() { { 0, 0 }, { 1, 0 }, { 2, 0 } };
                var p4S33ValueDic = row > 4 ? context.GetP4RowResult(33, row) : new Dictionary<int, object>() { { 0, 0 }, { 1, 0 }, { 2, 0 } };

                var p4S32ValueDic1 = row - 1 > 4 ? context.GetP4RowResult(32, row - 1) : new Dictionary<int, object>() { { 0, 0 }, { 1, 0 }, { 2, 0 } };
                var p4S33ValueDic1 = row - 1 > 4 ? context.GetP4RowResult(33, row - 1) : new Dictionary<int, object>() { { 0, 0 }, { 1, 0 }, { 2, 0 } };

                var valueDic = new int[4];
                var valueDic1 = new int[4];

                valueDic[0] = (int)p4S32ValueDic[0];
                valueDic[1] = (int)p4S32ValueDic[1];
                valueDic[2] = (int)p4S33ValueDic[0];
                valueDic[3] = (int)p4S33ValueDic[1];

                valueDic1[0] = (int)p4S32ValueDic1[0];
                valueDic1[1] = (int)p4S32ValueDic1[1];
                valueDic1[2] = (int)p4S33ValueDic1[0];
                valueDic1[3] = (int)p4S33ValueDic1[1];

                var mark = 0;
                for (int i = 0; i < 4; i++)
                {
                    if (P4S32Yellow.Contains(valueDic[i]) && P4S32Yellow.Contains(valueDic1[i]))
                    {
                        mark++;
                    }
                    else if (valueDic[i] > 8 && valueDic1[i] > 8)
                    {
                        mark++;
                    }
                    else if (valueDic[i] < 6 && valueDic1[i] < 6)
                    {
                        mark++;
                    }
                }
                context.SetP4Value(37, row, 0, mark);
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
