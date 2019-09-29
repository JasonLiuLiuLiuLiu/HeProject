using HeProject.Model;
using HeProject.ProgressHandler.P1;

namespace HeProject
{
    public class S1Handler : IP1Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            new P1HandleCommon().Hnalder(1, row, context);
            return null;
        }
    }
}
