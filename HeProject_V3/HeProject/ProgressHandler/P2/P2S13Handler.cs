using System;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S13Handler : IP2Handler
    {
        public string Handler(int row, ProcessContext context)
        {
            var source = new int[6][];
            for (var i = 0; i < 6; i++)
            {
                if (context.GetP1StepState(i + 7, row))
                {
                    source[i] = context.GetP1RowResult(i + 7, row).Select(u => (int)u.Value).ToArray();
                }
            }

            for (var i = 0; i < 6; i++)
            {
                if (source[i] == null || source[i].Length != 3) continue;

                if (source[i][1] < 5)
                {
                    context.SetP2Value(13, row, source[i][2] < 5 ? 0 : 1, true);
                }
                else
                {
                    context.SetP2Value(13, row, source[i][2] < 5 ? 2 : 3, true);
                }
            }

            return null;
        }
    }
}
