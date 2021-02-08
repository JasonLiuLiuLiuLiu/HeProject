using System;
using System.Collections.Generic;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S42P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            try
            {
                var sum1 = context.GetP4RowResult(19, row).CheckIntValue();
                var sum2 = context.GetP4RowResult(20, row).CheckIntValue();
                var sum3 = context.GetP4RowResult(21, row).CheckIntValue();
                for (int i = 0; i < 6; i++)
                {
                    context.SetP4Value(42, row, i, (int)sum1[i] + (int)sum2[i] + (int)sum3[i]);
                }
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
