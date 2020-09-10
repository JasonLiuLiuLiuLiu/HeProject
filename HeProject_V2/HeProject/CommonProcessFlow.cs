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
            return context;
        }
    }
}
