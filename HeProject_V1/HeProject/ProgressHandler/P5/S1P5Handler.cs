using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P5
{
    public class S1P5Handler : IP5Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var z0 =
                context.GetP1Value<int>(7, row, 0) + context.GetP1Value<int>(8, row, 0) +
                context.GetP2Value<int>(7, row, 0) + context.GetP2Value<int>(8, row, 0) +
                context.GetP3Value<int>(7, row, 0) + context.GetP3Value<int>(8, row, 0) +
                context.GetP4Value<int>(7, row, 0) + context.GetP4Value<int>(8, row, 0);

            var z1 =
                context.GetP1Value<int>(7, row, 1) + context.GetP1Value<int>(8, row, 1) +
                context.GetP2Value<int>(7, row, 1) + context.GetP2Value<int>(8, row, 1) +
                context.GetP3Value<int>(7, row, 1) + context.GetP3Value<int>(8, row, 1) +
                context.GetP4Value<int>(7, row, 1) + context.GetP4Value<int>(8, row, 1);

            var z2 =
                context.GetP1Value<int>(7, row, 2) + context.GetP1Value<int>(8, row, 2) +
                context.GetP2Value<int>(7, row, 2) + context.GetP2Value<int>(8, row, 2) +
                context.GetP3Value<int>(7, row, 2) + context.GetP3Value<int>(8, row, 2) +
                context.GetP4Value<int>(7, row, 2) + context.GetP4Value<int>(8, row, 2);

            context.SetP5Value(1, row, 0, z0);
            context.SetP5Value(1, row, 1, z1);
            context.SetP5Value(1, row, 2, z2);
            context.SetP5StepState(1, row, true);

            return null;
        }
    }
}
