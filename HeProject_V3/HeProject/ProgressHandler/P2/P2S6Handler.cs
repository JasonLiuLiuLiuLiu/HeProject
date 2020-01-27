using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S6Handler:IP2Handler
    {
        //百位加个位
        public string Handler(int stage, int row, ProcessContext context)
        {
            return new P2HandleCommon().GetOrder(stage, 6, row, context);
        }
    }
}
