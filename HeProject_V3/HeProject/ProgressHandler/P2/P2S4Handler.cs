using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S4Handler : IP2Handler
    {
        //百位加十位
        public string Handler(int row, ProcessContext context)
        {
            var source = new int[6][];
            for (var i = 0; i < 6; i++)
            {
                if (context.GetP1StepState(i + 7, row))
                {
                    source[i] = context.GetP1RowResult(i + 7, row).Select(u => (int)u.Value).ToArray();
                }
            }

            for (var i = 0; i < 6; i++)
            {
                if (source[i] == null || source[i].Length != 3) continue;

                if (source[i][1] < 5)
                {
                    context.SetP2Value(4, row, source[i][2] % 2 == 1 ? 0 : 1, true);
                }
                else
                {
                    context.SetP2Value(4, row, source[i][2] % 2 == 1 ? 2 : 3, true);
                }
            }

            return null;
        }
    }
}
