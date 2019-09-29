using System.Collections.Generic;
using System.Linq;
using HeProject.Model;
using HeProject.ProgressHandler.P1;
using NPOI.HSSF.Record;

namespace HeProject
{
    public class S3Handler : IP1Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            new P1HandleCommon().Hnalder(3, row, context);
            return null;
        }
    }
}
