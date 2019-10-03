using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    class S2P2Handler : IP2Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            new P2HandleCommon().GetOrder(2, row, context);
            return null;
        }
    }
}
