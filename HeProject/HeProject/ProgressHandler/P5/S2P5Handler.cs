using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P5
{
    public class S2P5Handler : IP5Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            int re1 = 0;
            int leng2 = 0;

            if (row == 0)
            {
                context.SetP5StepState(2, row, true);
                return null;
            }

            var preRow = row - 1;

            #region P1
            if (context.GetP1Value<int>(7, preRow, 0) > 1)
                re1 += context.GetP1Value<int>(7, row, 0);
            else
                leng2 += context.GetP1Value<int>(7, row, 0);

            if (context.GetP1Value<int>(7, preRow, 1) > 1)
                re1 += context.GetP1Value<int>(7, row, 1);
            else
                leng2 += context.GetP1Value<int>(7, row, 1);

            if (context.GetP1Value<int>(7, preRow, 2) > 1)
                re1 += context.GetP1Value<int>(7, row, 2);
            else
                leng2 += context.GetP1Value<int>(7, row, 2);

            if (context.GetP1Value<int>(8, preRow, 0) > 1)
                re1 += context.GetP1Value<int>(8, row, 0);
            else
                leng2 += context.GetP1Value<int>(8, row, 0);

            if (context.GetP1Value<int>(8, preRow, 1) > 1)
                re1 += context.GetP1Value<int>(8, row, 1);
            else
                leng2 += context.GetP1Value<int>(8, row, 1);

            if (context.GetP1Value<int>(8, preRow, 2) > 1)
                re1 += context.GetP1Value<int>(8, row, 2);
            else
                leng2 += context.GetP1Value<int>(8, row, 2);
            #endregion

            #region P2

            if (context.GetP2Value<int>(7, preRow, 0) > 1)
                re1 += context.GetP2Value<int>(7, row, 0);
            else
                leng2 += context.GetP2Value<int>(7, row, 0);

            if (context.GetP2Value<int>(7, preRow, 1) > 1)
                re1 += context.GetP2Value<int>(7, row, 1);
            else
                leng2 += context.GetP2Value<int>(7, row, 1);

            if (context.GetP2Value<int>(7, preRow, 2) > 1)
                re1 += context.GetP2Value<int>(7, row, 2);
            else
                leng2 += context.GetP2Value<int>(7, row, 2);

            if (context.GetP2Value<int>(8, preRow, 0) > 1)
                re1 += context.GetP2Value<int>(8, row, 0);
            else
                leng2 += context.GetP2Value<int>(8, row, 0);

            if (context.GetP2Value<int>(8, preRow, 1) > 1)
                re1 += context.GetP2Value<int>(8, row, 1);
            else
                leng2 += context.GetP2Value<int>(8, row, 1);

            if (context.GetP2Value<int>(8, preRow, 2) > 1)
                re1 += context.GetP2Value<int>(8, row, 2);
            else
                leng2 += context.GetP2Value<int>(8, row, 2);

            #endregion

            #region P3

            if (context.GetP3Value<int>(7, preRow, 0) > 1)
                re1 += context.GetP3Value<int>(7, row, 0);
            else
                leng2 += context.GetP3Value<int>(7, row, 0);

            if (context.GetP3Value<int>(7, preRow, 1) > 1)
                re1 += context.GetP3Value<int>(7, row, 1);
            else
                leng2 += context.GetP3Value<int>(7, row, 1);

            if (context.GetP3Value<int>(7, preRow, 2) > 1)
                re1 += context.GetP3Value<int>(7, row, 2);
            else
                leng2 += context.GetP3Value<int>(7, row, 2);

            if (context.GetP3Value<int>(8, preRow, 0) > 1)
                re1 += context.GetP3Value<int>(8, row, 0);
            else
                leng2 += context.GetP3Value<int>(8, row, 0);

            if (context.GetP3Value<int>(8, preRow, 1) > 1)
                re1 += context.GetP3Value<int>(8, row, 1);
            else
                leng2 += context.GetP3Value<int>(8, row, 1);

            if (context.GetP3Value<int>(8, preRow, 2) > 1)
                re1 += context.GetP3Value<int>(8, row, 2);
            else
                leng2 += context.GetP3Value<int>(8, row, 2);

            #endregion

            #region P4

            if (context.GetP4Value<int>(7, preRow, 0) > 1)
                re1 += context.GetP4Value<int>(7, row, 0);
            else
                leng2 += context.GetP4Value<int>(7, row, 0);

            if (context.GetP4Value<int>(7, preRow, 1) > 1)
                re1 += context.GetP4Value<int>(7, row, 1);
            else
                leng2 += context.GetP4Value<int>(7, row, 1);

            if (context.GetP4Value<int>(7, preRow, 2) > 1)
                re1 += context.GetP4Value<int>(7, row, 2);
            else
                leng2 += context.GetP4Value<int>(7, row, 2);

            if (context.GetP4Value<int>(8, preRow, 0) > 1)
                re1 += context.GetP4Value<int>(8, row, 0);
            else
                leng2 += context.GetP4Value<int>(8, row, 0);

            if (context.GetP4Value<int>(8, preRow, 1) > 1)
                re1 += context.GetP4Value<int>(8, row, 1);
            else
                leng2 += context.GetP4Value<int>(8, row, 1);

            if (context.GetP4Value<int>(8, preRow, 2) > 1)
                re1 += context.GetP4Value<int>(8, row, 2);
            else
                leng2 += context.GetP4Value<int>(8, row, 2);

            #endregion

            context.SetP5Value(2,row,0,re1);
            context.SetP5Value(2,row,1,leng2);
            context.SetP5StepState(2, row, true);

            return null;
        }
    }
}
