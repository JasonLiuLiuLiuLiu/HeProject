using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;
using HeProject.ProgressHandler.P3;

namespace HeProject.ProgressHandler.P4
{
    public class S1P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var a1Result = context.GetP1RowResult(1, row);
            var a2Result = context.GetP1RowResult(2, row);
            var a3Result = context.GetP1RowResult(3, row);

            for (int i = 0; i < StepLength.SourceLength; i++)
            {
                var count = 0;
                if (a1Result.ContainsKey(i) && (bool)a1Result[i])
                {
                    count++;
                }
                if (a2Result.ContainsKey(i) && (bool)a2Result[i])
                {
                    count++;
                }
                if (a3Result.ContainsKey(i) && (bool)a3Result[i])
                {
                    count++;
                }

                if (count != 0)
                {
                    context.SetP4Value(1, row, i, count);
                }
            }

            return null;
        }
    }
}
