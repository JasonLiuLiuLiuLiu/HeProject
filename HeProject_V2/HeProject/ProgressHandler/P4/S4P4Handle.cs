using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;
using HeProject.ProgressHandler.Common;
using HeProject.ProgressHandler.P3;

namespace HeProject.ProgressHandler.P4
{
    public class S4P4Handle:IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var a1Result = context.GetP1RowResult(1, row);
            var b1Result = context.GetP2RowResult(1, row);
            var c1Result = context.GetP3RowResult(1, row);

            for (int i = 0; i < StepLength.SourceLength; i++)
            {
                var count = 0;
                if (a1Result.ContainsKey(i) && (bool)a1Result[i])
                {
                    count++;
                }
                if (b1Result.ContainsKey(i) && (bool)b1Result[i])
                {
                    count++;
                }
                if (c1Result.ContainsKey(i) && (bool)c1Result[i])
                {
                    count++;
                }

                if (count != 0)
                {
                    context.SetP4Value(4, row, i, count);
                }
            }

            return null;
        }
    }
}
