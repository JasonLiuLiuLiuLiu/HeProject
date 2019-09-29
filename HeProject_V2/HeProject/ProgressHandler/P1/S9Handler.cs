using HeProject.Model;
using HeProject.ProgressHandler.P1;

namespace HeProject
{
    public class S9Handler : IP1Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            int c1 = 0;
            int c2 = 0;
            int c3 = 0;
            if (row == 0)
                return null;

            for (int i = 0; i < StepLength.P2; i++)
            {
                if(!context.GetP1Value<bool>(2,row,i))
                    continue;
                int distance = 0;
                for (int j = row-1; (j > row - 4) && j >= 0; j--)
                {
                    distance++;
                    if (context.GetP1Value<bool>(2, j, i))
                    {
                        break;
                    }
                }

                switch (distance)
                {
                    case 1:
                        c1++;
                        break;
                    case 2:
                        c2++;
                        break;
                    default:
                        c3++;
                        break;
                }
            }
            context.SetP1Value(9, row, 0, c1);
            context.SetP1Value(9, row, 1, c2);
            context.SetP1Value(9, row, 2, c3);
            return null;
        }
    }
}
