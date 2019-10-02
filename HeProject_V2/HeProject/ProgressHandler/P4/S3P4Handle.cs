using System;
using HeProject.Model;
using HeProject.ProgressHandler.P3;

namespace HeProject.ProgressHandler.P4
{
    public class S3P4Handle:IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var c1Result = context.GetP3RowResult(1, row);
            var c2Result = context.GetP3RowResult(2, row);
            var c3Result = context.GetP3RowResult(3, row);

            for (int i = 0; i < StepLength.SourceLength; i++)
            {
                var count = 0;
                if (c1Result.ContainsKey(i) && (bool)c1Result[i])
                {
                    count++;
                }
                if (c2Result.ContainsKey(i) && (bool)c2Result[i])
                {
                    count++;
                }
                if (c3Result.ContainsKey(i) && (bool)c3Result[i])
                {
                    count++;
                }

                if (count != 0)
                {
                    context.SetP4Value(3, row, i, count);
                }
            }

            return null;
        }
    }
}
