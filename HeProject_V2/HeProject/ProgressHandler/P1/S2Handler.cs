using HeProject.Model;
using HeProject.ProgressHandler.P1;
using NPOI.SS.Formula.Functions;

namespace HeProject
{
    public class S2Handler : IP1Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            if (row < 1)
                return null;
            var buffer = new int[2, StepLength.P2];
            for (int i = 0; i < StepLength.P2; i++)
            {
                buffer[0, i] = context.GetP1Value<int>(1, row - 1, i);
                buffer[1, i] = context.GetP1Value<int>(1, row, i);
            }

            for (int i = 0; i < StepLength.P2; i++)
            {
                if (isBig(buffer[0, i]) == isBig(buffer[1, i]))
                    context.SetP1Value(2, row, i, true);
            }

            return null;
        }

        private bool isBig(int value)
        {
            if (value > 2 && value < 6)
                return true;
            return false;
        }
    }
}
