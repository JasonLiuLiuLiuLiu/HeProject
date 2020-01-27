using System;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S4Handler:IP2Handler
    {
        //百位加十位
        public string Handler(int stage, int row, ProcessContext context)
        {
            var source = context.GetP1RowResult(stage + 7, row).Select(u => (int)u.Value).ToArray();
            if (source[0] < 5)
            {
                if (source[1] < 5)
                {

                }
                else
                {
                    
                }
            }
            else
            {
                if (source[1] < 5)
                {

                }
                else
                {

                }
            }

            return null;
        }
    }
}
