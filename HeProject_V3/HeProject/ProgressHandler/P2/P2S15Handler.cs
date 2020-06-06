using System;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S15Handler : IP2Handler
    {
        public string Handler(int row, ProcessContext context)
        {
            return new P2HandleCommon().GetOrder(15, row, context);
        }
    }
}
