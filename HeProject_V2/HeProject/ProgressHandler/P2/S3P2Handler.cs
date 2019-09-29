using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    class S3P2Handler : IP2Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            new P2HandleCommon().Hnalder(3, row, context);
            return null;
        }
    }
}
