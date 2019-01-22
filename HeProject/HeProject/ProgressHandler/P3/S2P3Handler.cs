using HeProject.Model;

namespace HeProject.ProgressHandler.P3
{
    public class S2P3Handler : IP3Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            context.SetP3StepState(1, row, true);

            if (row < 1)
                return null;
            var buffer = new int[2, StepLength.P2];
            for (int i = 0; i < StepLength.P2; i++)
            {
                buffer[0, i] = context.GetP3Value<int>(1, row - 1, i);
                buffer[1, i] = context.GetP3Value<int>(1, row, i);
            }

            for (int i = 0; i < StepLength.P2; i++)
            {
                if (isDouble(buffer[0, i]) == isDouble(buffer[1, i]))
                    context.SetP3Value(2, row, i, true);
            }
            return null;
        }
        private bool isDouble(int value)
        {
            if (value % 2 == 0)
                return true;
            return false;
        }
    }
}