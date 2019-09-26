using HeProject.Model;
using System;

namespace HeProject.ProgressHandler.P3
{
    public class S4P3Handler : IP3Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            try
            {
                if (row < 1)
                    return null;
                var p2Result = context.GetP3RowResult(2, row);
                var p3Result = context.GetP3RowResult(3, row - 1);
                bool[] result = new bool[StepLength.P4];
                for (int i = 0; i < StepLength.P4; i++)
                {
                    if ((bool)p2Result[i])
                        context.SetP3Value(4, row, (int)p3Result[i], true);
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