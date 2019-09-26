using HeProject.Model;

namespace HeProject.ProgressHandler.P3
{
    public class S6P3Handler : IP3Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            if (row < 1)
                return null;
            var p4Result = context.GetP3RowResult(4, row);
            var p5Result = context.GetP3RowResult(5, row - 1);
            bool[] result = new bool[StepLength.P6];
            for (int i = 0; i < StepLength.P6; i++)
            {
                if ((bool)p4Result[i])
                    context.SetP3Value(6, row, (int)p5Result[i], true);
            }
            return null;
        }
    }
}