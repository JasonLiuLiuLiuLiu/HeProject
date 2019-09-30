using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P3
{
    public class S0P3Handler : IP3Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            if (row <= 1)
                return null;
            var p10 = context.GetP1RowResult(0, row);
            var p11 = context.GetP1RowResult(0, row - 1);
            var p12 = context.GetP1RowResult(0, row - 2);
            var index = (p10.FirstOrDefault(u => (bool)u.Value).Key + p11.FirstOrDefault(u => (bool)u.Value).Key + p12.FirstOrDefault(u => (bool)u.Value).Key) %
                        StepLength.SourceLength;
            context.SetP3Value(0, row, index, true);

            return null;
        }
    }
}
