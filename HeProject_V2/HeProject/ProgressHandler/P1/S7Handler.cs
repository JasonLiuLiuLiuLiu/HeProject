using System.Linq;
using HeProject.Model;
using HeProject.ProgressHandler.P1;

namespace HeProject
{
    public class S7Handler : IP1Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var p4Result = context.GetP1RowResult(4, row);
            var c1 = p4Result.Where(u => u.Key < 3).Where(u => (bool) u.Value).Count();
            var c2 = p4Result.Where(u => u.Key > 2).Where(u => u.Key < 6).Where(u => (bool) u.Value).Count();
            var c3 = p4Result.Where(u => u.Key > 5).Where(u => (bool) u.Value).Count();
            context.SetP1Value(7,row,0,c1);
            context.SetP1Value(7,row,1,c2);
            context.SetP1Value(7,row,2,c3);
            return null;
        }
    }
}
