using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    class S1P2Handler : IP2Handler
    {
        private ProcessContext _context;

        public string Hnalder(int row, ProcessContext context)
        {
            new P2HandleCommon().Hnalder(1, row, context);
            return null;
        }
    }
}
