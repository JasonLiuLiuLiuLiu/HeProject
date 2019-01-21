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
        private Dictionary<int, Dictionary<int, Dictionary<int, object>>> valueP1Map;
        public readonly int Capacity;
        private const int stepCont = 9;
        private object _lock=new object();

        public void SetP1Value(int step, int row, int column, object value)
        {
            try
            {
                lock (_lock)
                {
                    if (!valueP1Map.ContainsKey(step))
                        valueP1Map.Add(step, new Dictionary<int, Dictionary<int, object>>(stepCont));
                    var stepValueMap = valueP1Map[step];
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

        public T GetP1Value<T>(int step, int row, int column)
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
                    if (!valueP1Map.ContainsKey(step))
                        valueP1Map.Add(step, new Dictionary<int, Dictionary<int, object>>(stepCont));
                    var stepValueMap = valueP1Map[step];
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

        public Dictionary<int, object> GetP1RowResult(int step, int row)
        {
            int loop = 0;
            while (!_stepState[step - 1, row])
            {
                if (loop > 100)
                    return new Dictionary<int, object>(Capacity);
                Thread.Sleep(50);
            }
            if (!valueP1Map.ContainsKey(step))
                valueP1Map.Add(step, new Dictionary<int, Dictionary<int, object>>(stepCont));
            var stepValueMap = valueP1Map[step];
            if (!stepValueMap.ContainsKey(row))
                stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
            return stepValueMap[row];
        }

        public void SetP1StepState(int step, int row, bool value)
        {
            _stepState[step - 1, row] = value;
        }

        public ProcessContext(int capacity)
        {
            Capacity = capacity;
            _stepState = new bool[stepCont, capacity];
            valueP1Map = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
        }
    }
}
