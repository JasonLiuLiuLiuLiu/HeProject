using HeProject.Model;
using HeProject.ProgressHandler.Common;

namespace HeProject
{
    public class S10Handler : IP1Handler
    {
        private ProcessContext context;

        public string Hnalder(int row, ProcessContext context)
        {
            if (row < 3)
                return null;
            this.context = context;

            context.SetP1Value(10, row, 0, GetInputs(row, 7, 0).CalculateSameLength());
            context.SetP1Value(10, row, 1, GetInputs(row, 7, 1).CalculateSameLength());
            context.SetP1Value(10, row, 2, GetInputs(row, 7, 2).CalculateSameLength());
            context.SetP1Value(10, row, 3, GetInputs(row, 8, 0).CalculateSameLength());
            context.SetP1Value(10, row, 4, GetInputs(row, 8, 1).CalculateSameLength());
            context.SetP1Value(10, row, 5, GetInputs(row, 8, 2).CalculateSameLength());

            return null;
        }

        private int[] GetInputs(int row, int step, int column)
        {
            var result = new int[4];
            result[0] = context.GetP1Value<int>(step, row, column);
            result[1] = context.GetP1Value<int>(step, row - 1, column);
            result[2] = context.GetP1Value<int>(step, row - 2, column);
            result[3] = context.GetP1Value<int>(step, row - 3, column);
            return result;
        }
    }
}
