using HeProject.Model;

namespace HeProject.ProgressHandler.P3
{
    public class S1P3Handler : IP3Handler
    {
        private ProcessContext _context;

        public string Hnalder(int row, ProcessContext context)
        {
            new P3HandleCommon().Hnalder(1, row, context);
            return null;
        }
    }
}