using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    class S0P2Handler : IP2Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            if (row == 0)
                return null;
            var p10 = context.GetP1RowResult(0, row);
            var p11 = context.GetP1RowResult(0, row - 1);
            var index = (p10.FirstOrDefault(u => (bool)u.Value).Key + p11.FirstOrDefault(u => (bool)u.Value).Key) %
                        StepLength.SourceLength;
            context.SetP2Value(0, row, index, true);

            return null;
        }
    }
}
