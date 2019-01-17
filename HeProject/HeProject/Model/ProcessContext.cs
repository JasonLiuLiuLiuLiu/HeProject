using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;

namespace HeProject.Model
{
    public class ProcessContext
    {
      
        private bool[,] _stepState;

        private Dictionary<int, Dictionary<int, Dictionary<int, object>>> valueMap;
        private readonly int _capacity;
        private const int stepCont = 9;

        public void SetValue(int step, int row, int column, object value)
        {
            if (!valueMap.ContainsKey(step))
                valueMap.Add(step, new Dictionary<int, Dictionary<int, object>>(stepCont));
            var stepValueMap = valueMap[step];
            if (!stepValueMap.ContainsKey(row))
                stepValueMap.Add(row, new Dictionary<int, object>(_capacity));
            var rowValueMap = stepValueMap[row];
            if (!rowValueMap.ContainsKey(column))
                rowValueMap.Add(column, null);

            rowValueMap[column] = value;
        }

        public T GetValue<T>(int step, int row, int column)
        {
            int loop = 0;
            while (!_stepState[step-1, row])
            {
                if (loop > 100)
                    return default(T);
                Thread.Sleep(50);
            }
            if (!valueMap.ContainsKey(step))
                valueMap.Add(step, new Dictionary<int, Dictionary<int, object>>(stepCont));
            var stepValueMap = valueMap[step];
            if (!stepValueMap.ContainsKey(row))
                stepValueMap.Add(row, new Dictionary<int, object>(_capacity));
            var rowValueMap = stepValueMap[row];
            if (!rowValueMap.ContainsKey(column))
                rowValueMap.Add(column, default(T));
            return (T)rowValueMap[column];
        }

        public void SetStepState(int step, int row,bool value)
        {
            _stepState[step - 1, row] = value;
        }

        public ProcessContext(int capacity)
        {
            _capacity = capacity;
            _stepState = new bool[stepCont, capacity];
            valueMap = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
        }
    }
}
