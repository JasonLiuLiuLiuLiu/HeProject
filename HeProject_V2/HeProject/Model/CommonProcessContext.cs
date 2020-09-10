using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;

namespace HeProject.Model
{
    public class CommonProcessContext
    {
        public CommonProcessContext(Dictionary<int, ProcessContext> resultDic,IWorkbook workbook)
        {
            this.ResultDic = resultDic;
            this.Workbook = workbook;
        }
        public Dictionary<int, ProcessContext> ResultDic { get; }
        public IWorkbook Workbook { get; }
        public int Shao { get; set; }
        public int Gu { get; set; }
    }
}
