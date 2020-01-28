using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P4
{
    public class S30P4Handle : IP4Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            return new P4MutilValueOrderHandle().GetOrder(30, row, context);
        }
    }
}
