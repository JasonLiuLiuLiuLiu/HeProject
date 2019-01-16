using HeProject.Model;

namespace HeProject
{
    public interface IProgressHandler
    {
        string Hnalder(int row, ProcessContext context);
    }
}