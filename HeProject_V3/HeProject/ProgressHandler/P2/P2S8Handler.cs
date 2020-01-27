using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S8Handler:IP2Handler
    {
        private static readonly int[] Ying = new[] { 1, 2, 5, 8, 9 };
        public string Handler(int stage, int row, ProcessContext context)
        {
            return new P2HandleCommon().GetOrder(stage, 8, row, context);
        }
    }
}
