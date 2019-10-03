using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;
using HeProject.ProgressHandler.Common;
using HeProject.ProgressHandler.P3;

namespace HeProject.ProgressHandler.P4
{
    public class S6P4Handle:IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var a3Result = context.GetP1RowResult(3, row);
            var b3Result = context.GetP2RowResult(3, row);
            var c3Result = context.GetP3RowResult(3, row);

            for (int i = 0; i < StepLength.SourceLength; i++)
            {
                var count = 0;
                if (a3Result.ContainsKey(i) && (bool)a3Result[i])
                {
                    count++;
                }
                if (b3Result.ContainsKey(i) && (bool)b3Result[i])
                {
                    count++;
                }
                if (c3Result.ContainsKey(i) && (bool)c3Result[i])
                {
                    count++;
                }

                if (count != 0)
                {
                    context.SetP4Value(6, row, i, count);
                }
            }

            return null;
        }
    }
}
