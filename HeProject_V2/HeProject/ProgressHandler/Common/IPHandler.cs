using HeProject.Model;

namespace HeProject.ProgressHandler.Common
{
    public interface IPHandler
    {
        string Hnalder(int row, ProcessContext context);
    }
}