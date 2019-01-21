using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;

namespace HeProject.Model
{
    public class ProcessContext
    {

        private bool[,] _stepP1State;
        private bool[,] _stepP2State;
        private bool[,] _stepP3State;
        private bool[,] _stepP4State;

        private Dictionary<int, Dictionary<int, Dictionary<int, object>>> valueP1Map;
        private Dictionary<int, Dictionary<int, Dictionary<int, object>>> valueP2Map;
        private Dictionary<int, Dictionary<int, Dictionary<int, object>>> valueP3Map;
        private Dictionary<int, Dictionary<int, Dictionary<int, object>>> valueP4Map;
        public readonly int Capacity;
        private const int stepCont = 9;
        private object _lockP1=new object();
        private object _lockP2=new object();
        private object _lockP3=new object();
        private object _lockP4=new object();

        #region P1

        public void SetP1Value(int step, int row, int column, object value)
        {
            try
            {
                lock (_lockP1)
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
                lock (_lockP1)
                {
                    int loop = 0;
                    while (!_stepP1State[step - 1, row])
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
            while (!_stepP1State[step - 1, row])
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
            _stepP1State[step - 1, row] = value;
        }


        #endregion

        #region P2

        public void SetP2Value(int step, int row, int column, object value)
        {
            try
            {
                lock (_lockP2)
                {
                    if (!valueP2Map.ContainsKey(step))
                        valueP2Map.Add(step, new Dictionary<int, Dictionary<int, object>>(stepCont));
                    var stepValueMap = valueP2Map[step];
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

        public T GetP2Value<T>(int step, int row, int column)
        {
            try
            {
                lock (_lockP2)
                {
                    int loop = 0;
                    while (!_stepP2State[step - 1, row])
                    {
                        if (loop > 100)
                            return default(T);
                        Thread.Sleep(50);
                    }
                    if (!valueP2Map.ContainsKey(step))
                        valueP2Map.Add(step, new Dictionary<int, Dictionary<int, object>>(stepCont));
                    var stepValueMap = valueP2Map[step];
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

        public Dictionary<int, object> GetP2RowResult(int step, int row)
        {
            int loop = 0;
            while (!_stepP2State[step - 1, row])
            {
                if (loop > 100)
                    return new Dictionary<int, object>(Capacity);
                Thread.Sleep(50);
            }
            if (!valueP2Map.ContainsKey(step))
                valueP2Map.Add(step, new Dictionary<int, Dictionary<int, object>>(stepCont));
            var stepValueMap = valueP2Map[step];
            if (!stepValueMap.ContainsKey(row))
                stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
            return stepValueMap[row];
        }

        public void SetP2StepState(int step, int row, bool value)
        {
            _stepP2State[step - 1, row] = value;
        }


        #endregion
        public ProcessContext(int capacity)
        {
            Capacity = capacity;
            _stepP1State = new bool[stepCont, capacity];
            _stepP2State = new bool[stepCont, capacity];
            valueP1Map = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
            valueP2Map = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
            valueP3Map = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
            valueP4Map = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
        }
    }
}
