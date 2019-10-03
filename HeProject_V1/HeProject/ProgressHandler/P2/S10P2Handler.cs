using HeProject.Model;
using HeProject.Part2;
using HeProject.ProgressHandler.Common;

namespace HeProject.ProgressHandler.P2
{
    class S10P2Handler : IP2Handler
    {
        private ProcessContext _context;

        public string Hnalder(int row, ProcessContext context)
        {
            if (row < 3)
                return null;
            _context = context;

            context.SetP2Value(10, row, 0, GetInputs(row, 7, 0).CalculateSameLength());
            context.SetP2Value(10, row, 1, GetInputs(row, 7, 1).CalculateSameLength());
            context.SetP2Value(10, row, 2, GetInputs(row, 7, 2).CalculateSameLength());
            context.SetP2Value(10, row, 3, GetInputs(row, 8, 0).CalculateSameLength());
            context.SetP2Value(10, row, 4, GetInputs(row, 8, 1).CalculateSameLength());
            context.SetP2Value(10, row, 5, GetInputs(row, 8, 2).CalculateSameLength());

            return null;
        }

        private int[] GetInputs(int row, int step, int column)
        {
            var result = new int[4];
            result[0] = _context.GetP2Value<int>(step, row, column);
            result[1] = _context.GetP2Value<int>(step, row - 1, column);
            result[2] = _context.GetP2Value<int>(step, row - 2, column);
            result[3] = _context.GetP2Value<int>(step, row - 3, column);
            return result;
        }
    }
}
