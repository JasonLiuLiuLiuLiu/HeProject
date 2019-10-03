using HeProject.Model;

namespace HeProject.ProgressHandler.P3
{
    public class S3P3Handler : IP3Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            new P3HandleCommon().GetOrder(3, row, context);
            return null;
        }
    }
}