using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S2Handler : IP2Handler
    {
        //十位加个位
        public string Handler(int row, ProcessContext context)
        {
            return new P2HandleCommon().GetOrder(2, row, context);
        }
    }
}
