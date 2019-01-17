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
        public readonly int Capacity;
        private const int stepCont = 9;
        private object _lock=new object();

        public void SetValue(int step, int row, int column, object value)
        {
            try
            {
                lock (_lock)
                {
                    if (!valueMap.ContainsKey(step))
                        valueMap.Add(step, new Dictionary<int, Dictionary<int, object>>(stepCont));
                    var stepValueMap = valueMap[step];
                    if (!stepValueMap.ContainsKey(row))
                        stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
                    var rowValueMap = stepValueMap[row];
                    if (!rowValueMap.ContainsKey(column))
                        rowValueMap.Add(column, null);

                    rowValueMap[column] = value;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"step:{step},row:{row},column:{column},value:{value},exception:{e.ToString()}");
                throw;
            }
          
        }

        public T GetValue<T>(int step, int row, int column)
        {
            try
            {
                lock (_lock)
                {
                    int loop = 0;
                    while (!_stepState[step - 1, row])
                    {
                        if (loop > 100)
                            return default(T);
                        Thread.Sleep(50);
                    }
                    if (!valueMap.ContainsKey(step))
                        valueMap.Add(step, new Dictionary<int, Dictionary<int, object>>(stepCont));
                    var stepValueMap = valueMap[step];
                    if (!stepValueMap.ContainsKey(row))
                        stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
                    var rowValueMap = stepValueMap[row];
                    if (!rowValueMap.ContainsKey(column))
                        rowValueMap.Add(column, default(T));
                    return (T)rowValueMap[column];
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"step:{step},row:{row},column:{column},exception:{e.ToString()}");
                throw;
            }
          
        }

        public Dictionary<int, object> GetRowResult(int step, int row)
        {
            int loop = 0;
            while (!_stepState[step - 1, row])
            {
                if (loop > 100)
                    return new Dictionary<int, object>(Capacity);
                Thread.Sleep(50);
            }
            if (!valueMap.ContainsKey(step))
                valueMap.Add(step, new Dictionary<int, Dictionary<int, object>>(stepCont));
            var stepValueMap = valueMap[step];
            if (!stepValueMap.ContainsKey(row))
                stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
            return stepValueMap[row];
        }

        public void SetStepState(int step, int row, bool value)
        {
            _stepState[step - 1, row] = value;
        }

        public ProcessContext(int capacity)
        {
            Capacity = capacity;
            _stepState = new bool[stepCont, capacity];
            valueMap = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
        }
    }
}
