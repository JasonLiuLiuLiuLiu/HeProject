using System;
using System.Collections.Generic;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S44P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            try
            {
                if (row < 2)
                    return null;
                var p382 = context.GetP4RowResult(42, row - 2).CheckIntValue();
                var p381 = context.GetP4RowResult(42, row - 1).CheckIntValue();
                var p380 = context.GetP4RowResult(42, row).CheckIntValue();
                var mark = 0;
                for (int i = 0; i < 6; i++)
                {
                    if ((int)p381[i] >= 3 && (int)p380[i] >= 3)
                    {
                        mark++;
                    }
                    if((int)p381[i] == 2&& (int)p382[i] >= 2&& (int)p380[i] >= 3)
                    {
                        mark++;
                    }
                }
                context.SetP4Value(44, row, 0, mark);
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
