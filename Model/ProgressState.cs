using System;
using System.Collections.Generic;
using System.Text;

namespace HeProject.Model
{
    public class ProgressState
    {
        public ProgressState(int step, int row)
        {
            Step = step;
            Row = row;
        }

        public int Step { get; set; }
        //-1失败,0成功,>0 行数
        public int Row { get; set; }
        public string ErrorMessage { get; set; }
    }
}
