using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S29P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            var valueDic = context.GetP4RowResult(27, row);
            for (int i = 0; i < 6; i++)
            {
                if (valueDic.ContainsKey(i) && (int)valueDic[i] >= 4)
                    context.SetP4Value(29, row, i, true);
            }
            return null;
        }
    }
}
