using HeProject.Model;

namespace HeProject
{
    public class P9Handler : IProgressHandler
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
                if(!context.GetValue<bool>(2,row,i))
                    continue;
                int distance = 0;
                for (int j = row-1; (j > row - 4) && j >= 0; j--)
                {
                    distance++;
                    if (context.GetValue<bool>(2, j, i))
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
            context.SetValue(9, row, 0, c1);
            context.SetValue(9, row, 1, c2);
            context.SetValue(9, row, 2, c3);
            return null;
        }
    }
}
