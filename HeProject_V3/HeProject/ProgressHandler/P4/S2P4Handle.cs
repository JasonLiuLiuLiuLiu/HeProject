using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;
using HeProject.ProgressHandler.Common;
using HeProject.ProgressHandler.P3;

namespace HeProject.ProgressHandler.P4
{
    public class S2P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var b1Result = context.GetP2RowResult(1, row);
            var b2Result = context.GetP2RowResult(2, row);
            var b3Result = context.GetP2RowResult(3, row);

            for (int i = 0; i < StepLength.SourceLength; i++)
            {
                var count = 0;
                if (b1Result.ContainsKey(i) && (bool)b1Result[i])
                {
                    count++;
                }
                if (b2Result.ContainsKey(i) && (bool)b2Result[i])
                {
                    count++;
                }
                if (b3Result.ContainsKey(i) && (bool)b3Result[i])
                {
                    count++;
                }

                if (count != 0)
                {
                    context.SetP4Value(2, row, i, count);
                }
            }

            return null;
        }
    }
}
