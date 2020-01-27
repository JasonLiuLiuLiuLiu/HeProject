using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S2Handler : IP2Handler
    {
        //十位加个位
        public string Handler(int stage, int row, ProcessContext context)
        {
            return new P2HandleCommon().GetOrder(stage, 2, row, context);
        }
    }
}
