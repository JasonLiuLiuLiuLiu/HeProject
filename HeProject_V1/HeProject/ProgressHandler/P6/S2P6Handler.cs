using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P6
{
    public class S2P6Handler : IP6Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            if (row < 2)
            {
                context.SetP6StepState(2, row, true);
                return null;
            }

            var pre = context.GetP6RowResult(1, row - 1);
            var current = context.GetP6RowResult(1, row);
            var result = new List<int>();

            for (int i = 0; i < pre.Count - 1; i++)
            {
                if ((int)pre[i] == (int)current[i] && (int)pre[i] == 1)
                    result.Add(1);
                else
                    result.Add(0);
            }
            context.SetP6Value(2, row, 0, result.Sum());

            context.SetP6StepState(2, row, true);
            return null;
        }
    }
}
