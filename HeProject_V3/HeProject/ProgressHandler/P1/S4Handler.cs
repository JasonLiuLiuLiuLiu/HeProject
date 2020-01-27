using System.Linq;
using HeProject.Model;
using HeProject.ProgressHandler.P1;

namespace HeProject
{
    public class S4Handler : IP1Handler
    {
        public string Handler(int row, ProcessContext context)
        {
            var s0 = context.GetP1RowResult(0, row);
            var s1 = context.GetP1RowResult(1, row);
            var s2 = context.GetP1RowResult(2, row);
            var s3 = context.GetP1RowResult(3, row);
            var s4 = context.GetP1RowResult(4, row);
            var s5 = context.GetP1RowResult(5, row);
            var s6 = context.GetP1RowResult(6, row);
            var stepIndex = 7;
            foreach (var keyValuePair in s0.Where(u => (bool)u.Value))
            {
                var s1Index = keyValuePair.Key;
                var value = new int[3];
                value[0] = (int)s4[s1Index];
                value[1] = (int)s5[value[0]];
                value[2] = (int)s6[value[1]];
                context.SetP1Value(stepIndex, row, 0, value[0]);
                context.SetP1Value(stepIndex, row, 1, value[1]);
                context.SetP1Value(stepIndex, row, 2, value[2]);
                context.SetP1StepState(stepIndex, row, true);
                stepIndex++;
            }

            return null;
        }
    }
}
