using System;
using System.Collections.Generic;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S51P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            try
            {
                if (row < 2)
                    return null;
                var v1 = context.GetP4Value<int>(40, row, 0);
                var v2 = context.GetP4Value<int>(44, row, 0);
                var v3 = context.GetP4Value<int>(48, row, 0);
                context.SetP4Value(51, row, 0, v1 + v2 + v3);
                return null;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}
