using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    class S4P4Handler : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            try
            {
                if (row < 1)
                    return null;
                var p2Result = context.GetP4RowResult(2, row);
                var p3Result = context.GetP4RowResult(3, row - 1);
                bool[] result = new bool[StepLength.P4];
                for (int i = 0; i < StepLength.P4; i++)
                {
                    if ((bool)p2Result[i])
                        context.SetP4Value(4, row, (int)p3Result[i], true);
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
