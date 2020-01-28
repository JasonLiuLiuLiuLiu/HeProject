using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S28Handler : IP2Handler
    {
        public string Handler(int row, ProcessContext context)
        {
            var result = new int[4];
            for (var i = 2; i < 10; i += 3)
            {
                var values = context.GetP2RowResult(i, row);
                for (int j = 0; j < 4; j++)
                {
                    if (values.ContainsKey(j) && (bool)values[j])
                    {
                        result[j]++;
                    }
                }
            }
            for (int i = 0; i < 4; i++)
            {
                context.SetP2Value(28, row, i, result[i]);
            }
            return null;
        }
    }
}
