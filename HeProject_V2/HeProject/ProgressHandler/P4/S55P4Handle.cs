using System;
using System.Collections.Generic;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S55P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            try
            {
                if (row < 2)
                    return null;
                var t1 = context.GetP4RowResult(38, row);
                var t2 = context.GetP4RowResult(42, row);
                var t3 = context.GetP4RowResult(46, row);
                var p1 = context.GetP4RowResult(38, row - 1);
                var p2 = context.GetP4RowResult(42, row - 1);
                var p3 = context.GetP4RowResult(46, row - 1);
                var sum = 0;
                for (int i = 0; i < 6; i++)
                {
                    if ((p1.ContainsKey(i) && (int)p1[i] == 0 || !p1.ContainsKey(i)) && t1.ContainsKey(i))
                    {
                        sum += (int)t1[i];
                    }
                    if ((p2.ContainsKey(i) && (int)p2[i] == 0 || !p2.ContainsKey(i)) && t2.ContainsKey(i))
                    {
                        sum += (int)t2[i];
                    }
                    if ((p3.ContainsKey(i) && (int)p3[i] == 0 || !p3.ContainsKey(i)) && t3.ContainsKey(i))
                    {
                        sum += (int)t3[i];
                    }
                }
                context.SetP4Value(55, row, 0, sum);
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
