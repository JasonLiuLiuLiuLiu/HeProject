using System;
using System.Collections.Generic;
using System.Text;
using HeProject.Model;
using HeProject.Part2;

namespace HeProject.ProgressHandler.P2
{
    class S6P2Handler : IP2Handler
    {
        public string Hnalder(int row, ProcessContext context)
        {
            Console.WriteLine(row);
            return null;
        }
    }
}
