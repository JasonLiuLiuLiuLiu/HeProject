using System;
using System.Collections.Generic;
using System.Text;

namespace HeProject.Model
{
    public class ProgressState
    {
        public ProgressState(int step, int row,int pNum=0)
        {
            Step = step;
            Row = row;
            PNum = pNum;
        }

        public int Stage { get; set; }
        public int PNum { get; set; }

        public int Step { get; set; }
        //-1失败,-2成功,>0 行数
        public int Row { get; set; }
        public string ErrorMessage { get; set; }
    }
}
