using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.Common
{
    interface ICPHandler
    {
        CommonProcessContext Run(CommonProcessContext context);
    }
}
