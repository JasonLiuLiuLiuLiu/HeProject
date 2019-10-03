using System;
using HeProject.Model;
using HeProject.ProgressHandler.P1;
using NPOI.SS.Formula.Functions;

namespace HeProject
{
    public class S2Handler : IP1Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            new P1HandleCommon().GetOrder(2, row, context);
            return null;
        }
    }
}
