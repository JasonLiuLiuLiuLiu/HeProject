using HeProject.Model;

namespace HeProject.ProgressHandler.P3
{
    public class S2P3Handler : IP3Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            new P3HandleCommon().GetOrder(2, row, context);
            return null;
        }
    }
}