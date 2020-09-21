using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;
using HeProject.ProgressHandler.CP;

namespace HeProject
{
    public class CommonProcessFlow
    {
        public CommonProcessContext Process(CommonProcessContext context)
        {
            new ShaoGuHandler().Run(context);
            new CopyHandler().Run(context);
            new TenBingAndSmall().Run(context);
            return context;
        }
    }
}
