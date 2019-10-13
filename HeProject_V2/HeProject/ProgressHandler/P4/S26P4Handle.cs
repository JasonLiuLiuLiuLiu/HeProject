using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S26P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var p4S1 = context.GetP4RowResult(1, row);
            var p4S2 = context.GetP4RowResult(2, row);
            var p4S3 = context.GetP4RowResult(3, row);
            var p4S4 = context.GetP4RowResult(4, row);
            var p4S5 = context.GetP4RowResult(5, row);
            var p4S6 = context.GetP4RowResult(6, row);

            var result = new int[StepLength.SourceLength];
            for (int i = 0; i < StepLength.SourceLength; i++)
            {
                result[i] =
                    GetValueOrDefault(p4S1, i) +
                    GetValueOrDefault(p4S2, i) +
                    GetValueOrDefault(p4S3, i) +
                    GetValueOrDefault(p4S4, i) +
                    GetValueOrDefault(p4S5, i) +
                    GetValueOrDefault(p4S6, i);
            }

            for (int i = 0; i < StepLength.SourceLength; i++)
            {
                context.SetP4Value(26, row, i, result[i]);
            }
            return null;
        }

        private int GetValueOrDefault(Dictionary<int, object> values, int column)
        {
            if (values.ContainsKey(column))
                return (int)values[column];

            return 0;
        }
    }
}
