using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    class S8P2Handler : IP2Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var p6Result = context.GetP2RowResult(6, row);
            var c1 = p6Result.Where(u => u.Key < 3).Where(u => (bool)u.Value).Count();
            var c2 = p6Result.Where(u => u.Key > 2).Where(u => u.Key < 6).Where(u => (bool)u.Value).Count();
            var c3 = p6Result.Where(u => u.Key > 5).Where(u => (bool)u.Value).Count();
            context.SetP2Value(8, row, 0, c1);
            context.SetP2Value(8, row, 1, c2);
            context.SetP2Value(8, row, 2, c3);
            return null;
        }
    }
}
