using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S3Handler : IP2Handler
    {
        //百位加个位
        public string Handler(int stage, int row, ProcessContext context)
        {
            var source = context.GetP1RowResult(stage + 7, row).Select(u => (int)u.Value).ToArray();

            if (source[0] % 2 == 1)
            {
                if (source[2] % 2 == 1)
                {
                    context.SetP2Value(stage, 3, row, 0, true);
                }
                else
                {
                    context.SetP2Value(stage, 3, row, 1, true);
                }
            }
            else
            {
                if (source[2] % 2 == 1)
                {
                    context.SetP2Value(stage, 3, row, 2, true);
                }
                else
                {
                    context.SetP2Value(stage, 3, row, 3, true);
                }
            }

            return null;
        }
    }
}
