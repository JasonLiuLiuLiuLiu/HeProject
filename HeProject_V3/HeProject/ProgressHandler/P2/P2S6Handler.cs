﻿using HeProject.Model;

namespace HeProject.ProgressHandler.P2
{
    public class P2S6Handler : IP2Handler
    {
        //百位加个位
        public string Handler(int row, ProcessContext context)
        {
            return new P2HandleCommon().GetOrder(6, row, context);
        }
    }
}
