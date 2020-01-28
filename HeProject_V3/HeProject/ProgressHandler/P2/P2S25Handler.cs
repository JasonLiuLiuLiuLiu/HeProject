using System;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S25Handler : IP2Handler
    {
        private static readonly int[] Yin = new[] { 1, 2, 5, 8, 9 };
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

                if (Yin.Contains(source[i][0]))
                {
                    context.SetP2Value(25, row, Yin.Contains(source[i][2]) ? 0 : 1, true);
                }
                else
                {
                    context.SetP2Value(25, row, Yin.Contains(source[i][2]) ? 2 : 3, true);
                }
            }

            return null;
        }
    }
}
