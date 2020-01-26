using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S27P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var p4S19 = context.GetP4RowResult(19, row);
            var p4S20 = context.GetP4RowResult(20, row);
            var p4S21 = context.GetP4RowResult(21, row);
            var p4S22 = context.GetP4RowResult(22, row);
            var p4S23 = context.GetP4RowResult(23, row);
            var p4S24 = context.GetP4RowResult(24, row);

            var result = new int[StepLength.SourceLength];
            for (int i = 0; i < StepLength.SourceLength; i++)
            {
                result[i] =
                    GetValueOrDefault(p4S19, i) +
                    GetValueOrDefault(p4S20, i) +
                    GetValueOrDefault(p4S21, i) +
                    GetValueOrDefault(p4S22, i) +
                    GetValueOrDefault(p4S23, i) +
                    GetValueOrDefault(p4S24, i);
            }

            for (int i = 0; i < StepLength.SourceLength; i++)
            {
                context.SetP4Value(27, row, i, result[i]);
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
