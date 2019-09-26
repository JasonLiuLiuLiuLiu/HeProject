using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    class S7P4Handler : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var p4Result = context.GetP4RowResult(4, row);
            var c1 = p4Result.Where(u => u.Key < 3).Where(u => (bool)u.Value).Count();
            var c2 = p4Result.Where(u => u.Key > 2).Where(u => u.Key < 6).Where(u => (bool)u.Value).Count();
            var c3 = p4Result.Where(u => u.Key > 5).Where(u => (bool)u.Value).Count();
            context.SetP4Value(7, row, 0, c1);
            context.SetP4Value(7, row, 1, c2);
            context.SetP4Value(7, row, 2, c3);
            return null;
        }
    }
}
