using System.Linq;
using HeProject.Model;

namespace HeProject
{
    public class P7Handler : IProgressHandler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var p4Result = context.GetRowResult(4, row);
            var c1 = p4Result.Where(u => u.Key < 3).Where(u => (bool) u.Value).Count();
            var c2 = p4Result.Where(u => u.Key > 2).Where(u => u.Key < 6).Where(u => (bool) u.Value).Count();
            var c3 = p4Result.Where(u => u.Key > 5).Where(u => (bool) u.Value).Count();
            context.SetValue(7,row,0,c1);
            context.SetValue(7,row,1,c2);
            context.SetValue(7,row,2,c3);
            return null;
        }
    }
}
