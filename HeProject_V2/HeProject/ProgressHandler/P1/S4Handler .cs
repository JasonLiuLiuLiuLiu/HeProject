using System;
using System.Collections.Generic;
using HeProject.Model;
using HeProject.ProgressHandler.P1;

namespace HeProject
{
    public class S4Handler : IP1Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            try
            {
                if (row < 1)
                    return null;
                var p2Result = context.GetP1RowResult(2, row);
                var p3Result = context.GetP1RowResult(3, row - 1);
                bool[] result = new bool[StepLength.P4];
                for (int i = 0; i < StepLength.P4; i++)
                {
                    if ((bool)p2Result[i])
                        context.SetP1Value(4, row, (int)p3Result[i], true);
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
