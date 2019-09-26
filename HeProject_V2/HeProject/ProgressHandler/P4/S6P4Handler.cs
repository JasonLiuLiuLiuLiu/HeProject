using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    class S6P4Handler : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            if (row < 1)
                return null;
            var p4Result = context.GetP4RowResult(4, row);
            var p5Result = context.GetP4RowResult(5, row - 1);
            bool[] result = new bool[StepLength.P6];
            for (int i = 0; i < StepLength.P6; i++)
            {
                if ((bool)p4Result[i])
                    context.SetP4Value(6, row, (int)p5Result[i], true);
            }
            return null;
        }
    }
}
