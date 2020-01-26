using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;
using HeProject.ProgressHandler.Common;
using HeProject.ProgressHandler.P3;

namespace HeProject.ProgressHandler.P4
{
    public class S5P4Handle:IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var a2Result = context.GetP1RowResult(2, row);
            var b2Result = context.GetP2RowResult(2, row);
            var c2Result = context.GetP3RowResult(2, row);

            for (int i = 0; i < StepLength.SourceLength; i++)
            {
                var count = 0;
                if (a2Result.ContainsKey(i) && (bool)a2Result[i])
                {
                    count++;
                }
                if (b2Result.ContainsKey(i) && (bool)b2Result[i])
                {
                    count++;
                }
                if (c2Result.ContainsKey(i) && (bool)c2Result[i])
                {
                    count++;
                }

                if (count != 0)
                {
                    context.SetP4Value(5, row, i, count);
                }
            }

            return null;
        }
    }
}
