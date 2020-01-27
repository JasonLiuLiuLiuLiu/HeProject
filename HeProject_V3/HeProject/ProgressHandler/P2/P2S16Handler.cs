﻿using System;
using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S16Handler:IP2Handler
    {
        public string Handler(int stage, int row, ProcessContext context)
        {
            return new P2HandleCommon().GetOrder(stage, 16, row, context,12);
        }
    }
}
