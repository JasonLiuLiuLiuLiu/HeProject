using System;
using System.Collections.Generic;
using System.Text;

namespace HeProject.Model
{
    public class ProcessContext
    {
        public int[,] InputData { get; set; }
        public bool?[,] P2 { get; set; }

        public ProcessContext(int capacity)
        {
            InputData = new int[capacity, 6];
            P2 = new bool?[capacity, 6];
        }
    }
}
