using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;

namespace HeProject.ProgressHandler.P5
{
    public class S3P5Handler : IP5Handler
    {
        private int _p1, _p2, _p3;
        public string Hnalder(int row, ProcessContext context)
        {
            AddValue(context.GetP1Value<int>(10, row, 0));
            AddValue(context.GetP1Value<int>(10, row, 1));
            AddValue(context.GetP1Value<int>(10, row, 2));
            AddValue(context.GetP1Value<int>(10, row, 3));
            AddValue(context.GetP1Value<int>(10, row, 4));
            AddValue(context.GetP1Value<int>(10, row, 5));

            AddValue(context.GetP2Value<int>(10, row, 0));
            AddValue(context.GetP2Value<int>(10, row, 1));
            AddValue(context.GetP2Value<int>(10, row, 2));
            AddValue(context.GetP2Value<int>(10, row, 3));
            AddValue(context.GetP2Value<int>(10, row, 4));
            AddValue(context.GetP2Value<int>(10, row, 5));

            AddValue(context.GetP3Value<int>(10, row, 0));
            AddValue(context.GetP3Value<int>(10, row, 1));
            AddValue(context.GetP3Value<int>(10, row, 2));
            AddValue(context.GetP3Value<int>(10, row, 3));
            AddValue(context.GetP3Value<int>(10, row, 4));
            AddValue(context.GetP3Value<int>(10, row, 5));

            AddValue(context.GetP4Value<int>(10, row, 0));
            AddValue(context.GetP4Value<int>(10, row, 1));
            AddValue(context.GetP4Value<int>(10, row, 2));
            AddValue(context.GetP4Value<int>(10, row, 3));
            AddValue(context.GetP4Value<int>(10, row, 4));
            AddValue(context.GetP4Value<int>(10, row, 5));

            context.SetP5Value(3, row, 0, _p1);
            context.SetP5Value(3, row, 1, _p2);
            context.SetP5Value(3, row, 2, _p3);
            context.SetP5StepState(3, row, true);
            return null;
        }

        private void AddValue(int value)
        {
            switch (value)
            {
                case 1:
                    _p1++;
                    break;
                case 2:
                    _p2++;
                    break;
                case 3:
                    _p3++;
                    break;
                default:
                    return;
            }
        }
    }
}
