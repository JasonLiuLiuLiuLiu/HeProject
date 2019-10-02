using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S9P4Handle:IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var sum01 = 0;
            var sum23 = 0;
            var sum45 = 0;

            var aSum = context.GetP4RowResult(3, row);

            foreach (var sum in aSum)
            {
                switch (sum.Key)
                {
                    case 0:
                        sum01 += (int)sum.Value;
                        break;
                    case 1:
                        sum01 += (int)sum.Value;
                        break;
                    case 2:
                        sum23 += (int)sum.Value;
                        break;
                    case 3:
                        sum23 += (int)sum.Value;
                        break;
                    case 4:
                        sum45 += (int)sum.Value;
                        break;
                    case 5:
                        sum45 += (int)sum.Value;
                        break;
                }
            }

            context.SetP4Value(9, row, 0, sum01);
            context.SetP4Value(9, row, 1, sum23);
            context.SetP4Value(9, row, 2, sum45);

            return null;
        }
    }
}
