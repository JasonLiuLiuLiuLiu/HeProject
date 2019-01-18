using HeProject.Model;
using System.Linq;

namespace HeProject
{
    public class P8Handler : IProgressHandler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var p6Result = context.GetRowResult(6, row);
            var c1 = p6Result.Where(u => u.Key < 3).Where(u => (bool)u.Value).Count();
            var c2 = p6Result.Where(u => u.Key > 2).Where(u => u.Key < 6).Where(u => (bool)u.Value).Count();
            var c3 = p6Result.Where(u => u.Key > 5).Where(u => (bool)u.Value).Count();
            context.SetValue(8, row, 0, c1);
            context.SetValue(8, row, 1, c2);
            context.SetValue(8, row, 2, c3);
            return null;
        }
    }
}
