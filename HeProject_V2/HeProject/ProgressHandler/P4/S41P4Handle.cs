using System;
using System.Collections.Generic;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S41P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            try
            {
                if (row < 1)
                    return null;
                var p381 = context.GetP4RowResult(38, row - 1).CheckIntValue();
                var p380 = context.GetP4RowResult(38, row).CheckIntValue();
                var mark = 0;
                for (int i = 0; i < 6; i++)
                {
                    if ((int)p381[i] ==0 && (int)p380[i] == 0)
                    {
                        mark++;
                    }
                }
                context.SetP4Value(41, row, 0, mark);
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
