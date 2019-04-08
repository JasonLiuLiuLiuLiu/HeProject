using HeProject.Model;
using HeProject.ProgressHandler.Common;

namespace HeProject.ProgressHandler.P3
{
    public class S10P3Handler : IP3Handler
    {
        private ProcessContext _context;

        public string Hnalder(int row, ProcessContext context)
        {
            if (row < 3)
                return null;
            _context = context;

            context.SetP3Value(10, row, 0, GetInputs(row, 7, 0).CalculateSameLength());
            context.SetP3Value(10, row, 1, GetInputs(row, 7, 1).CalculateSameLength());
            context.SetP3Value(10, row, 2, GetInputs(row, 7, 2).CalculateSameLength());
            context.SetP3Value(10, row, 3, GetInputs(row, 8, 0).CalculateSameLength());
            context.SetP3Value(10, row, 4, GetInputs(row, 8, 1).CalculateSameLength());
            context.SetP3Value(10, row, 5, GetInputs(row, 8, 2).CalculateSameLength());

            return null;
        }

        private int[] GetInputs(int row, int step, int column)
        {
            var result = new int[4];
            result[0] = _context.GetP3Value<int>(step, row, column);
            result[1] = _context.GetP3Value<int>(step, row - 1, column);
            result[2] = _context.GetP3Value<int>(step, row - 2, column);
            result[3] = _context.GetP3Value<int>(step, row - 3, column);
            return result;
        }
    }
}