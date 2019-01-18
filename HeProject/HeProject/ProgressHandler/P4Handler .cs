using System;
using System.Collections.Generic;
using HeProject.Model;

namespace HeProject
{
    public class P4Handler : IProgressHandler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            try
            {
                if (row < 1)
                    return null;
                var p2Result = context.GetRowResult(2, row);
                var p3Result = context.GetRowResult(3, row - 1);
                bool[] result = new bool[StepLength.P4];
                for (int i = 0; i < StepLength.P4; i++)
                {
                    if ((bool)p2Result[i])
                        context.SetValue(4, row, (int)p3Result[i], true);
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
