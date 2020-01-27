using System;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S22Handler:IP2Handler
    {
        private static readonly int[] Ying = new[] { 1, 2, 5, 8, 9 };
        public string Handler(int stage, int row, ProcessContext context)
        {
            var source = context.GetP1RowResult(stage + 7, row).Select(u => (int)u.Value).ToArray();
            if (Ying.Contains(source[1]))
            {
                if (Ying.Contains(source[2]))
                {
                    context.SetP2Value(stage, 22, row, 0, true);
                }
                else
                {
                    context.SetP2Value(stage, 22, row, 1, true);
                }
            }
            else
            {
                if (Ying.Contains(source[2]))
                {
                    context.SetP2Value(stage, 22, row, 2, true);
                }
                else
                {
                    context.SetP2Value(stage, 22, row, 3, true);
                }
            }

            return null;
        }
    }
}
