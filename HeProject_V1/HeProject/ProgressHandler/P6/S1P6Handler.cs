using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P6
{
    public class S1P6Handler : IP6Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            if (row == 0)
            {
                context.SetP6StepState(1, row, true);
                return null;
            }

           
            var pre = new List<int>();
            var current = new List<int>();
            var result = new List<int>();
            pre.AddRange(context.GetP1RowResult(7, row - 1).Select(u => (int)u.Value));
            pre.AddRange(context.GetP1RowResult(8, row - 1).Select(u => (int)u.Value));
            pre.AddRange(context.GetP2RowResult(7, row - 1).Select(u => (int)u.Value));
            pre.AddRange(context.GetP2RowResult(8, row - 1).Select(u => (int)u.Value));
            pre.AddRange(context.GetP3RowResult(7, row - 1).Select(u => (int)u.Value));
            pre.AddRange(context.GetP3RowResult(8, row - 1).Select(u => (int)u.Value));
            pre.AddRange(context.GetP4RowResult(7, row - 1).Select(u => (int)u.Value));
            pre.AddRange(context.GetP4RowResult(8, row - 1).Select(u => (int)u.Value));


            current.AddRange(context.GetP1RowResult(7, row).Select(u => (int)u.Value));
            current.AddRange(context.GetP1RowResult(8, row).Select(u => (int)u.Value));
            current.AddRange(context.GetP2RowResult(7, row).Select(u => (int)u.Value));
            current.AddRange(context.GetP2RowResult(8, row).Select(u => (int)u.Value));
            current.AddRange(context.GetP3RowResult(7, row).Select(u => (int)u.Value));
            current.AddRange(context.GetP3RowResult(8, row).Select(u => (int)u.Value));
            current.AddRange(context.GetP4RowResult(7, row).Select(u => (int)u.Value));
            current.AddRange(context.GetP4RowResult(8, row).Select(u => (int)u.Value));

            for (var i = 0; i < pre.Count; i++)
            {
                if ((pre[i] == 3 && 3 == current[i]) || pre[i] == 0 && 0 == current[i])
                    result.Add(1);
                else
                    result.Add(0);
            }

            for (int i = 0; i < result.Count; i++)
            {
                context.SetP6Value(1, row, i, result[i]);
            }
            context.SetP6Value(1, row, result.Count, result.Sum());

            context.SetP6StepState(1, row, true);
            return null;
        }
    }
}
