using HeProject.Model;

namespace HeProject.ProgressHandler.Common
{
    public interface IPHandler
    {
        string Handler(int row, ProcessContext context);
    }
}