using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S31P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            return new P4MutilValueOrderHandle().GetOrder(31, row, context);
        }
    }
}
