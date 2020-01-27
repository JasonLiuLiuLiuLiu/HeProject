﻿using System;
using System.Linq;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S10Handler:IP2Handler
    {
        public string Handler(int stage, int row, ProcessContext context)
        {
            var source = context.GetP1RowResult(stage + 7, row).Select(u => (int)u.Value).ToArray();
            if (source[0] < 5)
            {
                if (source[1] < 5)
                {
                    context.SetP2Value(stage, 10, row, 0, true);
                }
                else
                {
                    context.SetP2Value(stage, 10, row, 1, true);
                }
            }
            else
            {
                if (source[1] < 5)
                {
                    context.SetP2Value(stage, 10, row, 2, true);
                }
                else
                {
                    context.SetP2Value(stage, 10, row, 3, true);
                }
            }

            return null;
        }
    }
}
