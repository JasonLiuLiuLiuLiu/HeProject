using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;
using HeProject.Part2;

namespace HeProject.ProgressHandler.P2
{
    class S2P2Handler : IP2Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {

            if (row < 3)
            {
                for (int i = 0; i < StepLength.P2; i++)
                {
                    context.SetP2Value(2, row, i, false);
                }

                return null;
            }
            for (int i = 0; i < StepLength.P2; i++)
            {
                var state = new bool[2];
                var valueRow = context.GetP1Value<bool>(2, row, i);
                var valueRow_1 = context.GetP1Value<bool>(2, row - 1, i);
                var valueRow_2 = context.GetP1Value<bool>(2, row - 2, i);
                state[0] = valueRow == valueRow_1;
                state[1] = valueRow_1 == valueRow_2;
                context.SetP2Value(2, row, i, state[0] == state[1]);
            }

            return null;
        }
    }
}
