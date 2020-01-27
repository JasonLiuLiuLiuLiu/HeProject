using HeProject.Model;
using HeProject.ProgressHandler.Common;

namespace HeProject.ProgressHandler.P2
{
    public interface IP2Handler
    {
        string Handler(int stage, int row, ProcessContext context);
    }
}