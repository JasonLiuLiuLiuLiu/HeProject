using System;
using System.Collections.Generic;
using System.Threading;

namespace HeProject.Model
{
    public class ProcessContext
    {
        private readonly bool[,] _stepP1State;
        private readonly Dictionary<int, bool[,]> _stepP2State;


        public readonly Dictionary<int, Dictionary<int, Dictionary<int, object>>> ValueP1Map;
        public readonly Dictionary<int, Dictionary<int, Dictionary<int, Dictionary<int, object>>>> ValueP2Map;
        public readonly int Capacity;
        private const int StepCont = 50;

        private readonly object _lockP1 = new object();
        private readonly object _lock = new object();

        #region P1

        public void SetP1Value(int step, int row, int column, object value)
        {
            try
            {
                lock (_lockP1)
                {
                    if (!ValueP1Map.ContainsKey(step))
                        ValueP1Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = ValueP1Map[step];
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
                Console.WriteLine($"step:{step},row:{row},column:{column},value:{value},exception:{e}");
                throw;
            }
        }

        public T GetP1Value<T>(int step, int row, int column)
        {
            try
            {
                int loop = 0;
                while (!_stepP1State[step, row])
                {
                    if (loop > 100)
                        return default(T);
                    Thread.Sleep(50);
                }
                lock (_lockP1)
                {
                    if (!ValueP1Map.ContainsKey(step))
                        ValueP1Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = ValueP1Map[step];
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
                Console.WriteLine($"step:{step},row:{row},column:{column},exception:{e}");
                throw;
            }
        }

        public Dictionary<int, object> GetP1RowResult(int step, int row)
        {
            int loop = 0;
            while (!_stepP1State[step, row])
            {
                if (loop > 100)
                    return new Dictionary<int, object>(Capacity);
                Thread.Sleep(50);
            }
            lock (_lockP1)
            {
                if (!ValueP1Map.ContainsKey(step))
                    ValueP1Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));

                var stepValueMap = ValueP1Map[step];
                if (!stepValueMap.ContainsKey(row))
                    stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
                var result = new Dictionary<int, object>();
                foreach (var item in stepValueMap[row])
                {
                    result.Add(item.Key, item.Value);
                }
                return result;
            }

        }

        public void SetP1StepState(int step, int row, bool value)
        {
            _stepP1State[step, row] = value;
        }

        public bool GetP1StepState(int step, int row)
        {
            return _stepP1State[step, row];
        }

        #endregion

        #region P2

        public void SetP2Value(int stage, int step, int row, int column, object value)
        {
            try
            {
                lock (_lock)
                {
                    if (!ValueP2Map.ContainsKey(stage))
                        ValueP2Map.Add(stage, new Dictionary<int, Dictionary<int, Dictionary<int, object>>>());
                    if (!ValueP2Map[stage].ContainsKey(step))
                        ValueP2Map[stage].Add(step, new Dictionary<int, Dictionary<int, object>>());
                    var stepValueMap = ValueP2Map[stage][step];
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
                Console.WriteLine($"stage:{stage},step:{step},row:{row},column:{column},value:{value},exception:{e}");
                throw;
            }
        }

        public T GetP2Value<T>(int stage, int step, int row, int column)
        {
            CheckP2StepStage(stage);
            try
            {
                int loop = 0;
                while (!_stepP2State[stage][step, row])
                {
                    if (loop > 100)
                        return default(T);
                    Thread.Sleep(50);
                }
                lock (_lock)
                {
                    if (!ValueP2Map.ContainsKey(stage))
                        ValueP2Map.Add(stage, new Dictionary<int, Dictionary<int, Dictionary<int, object>>>());
                    if (!ValueP2Map[stage].ContainsKey(step))
                        ValueP2Map[stage].Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = ValueP2Map[stage][step];
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
                Console.WriteLine($"stage:{stage},step:{step},row:{row},column:{column},exception:{e}");
                throw;
            }
        }

        private void CheckP2StepStage(int stage)
        {
            if (!_stepP2State.ContainsKey(stage))
            {
                _stepP2State.Add(stage, new bool[StepCont, Capacity]);
            }
        }

        public Dictionary<int, object> GetP2RowResult(int stage, int step, int row)
        {
            CheckP2StepStage(stage);
            int loop = 0;
            while (!_stepP2State[stage][step, row])
            {
                if (loop > 100)
                    return new Dictionary<int, object>(Capacity);
                Thread.Sleep(50);
            }
            lock (_lock)
            {
                if (!ValueP2Map.ContainsKey(stage))
                    ValueP2Map.Add(stage, new Dictionary<int, Dictionary<int, Dictionary<int, object>>>());

                if (!ValueP2Map[stage].ContainsKey(step))
                    ValueP2Map[stage].Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));

                var stepValueMap = ValueP2Map[stage][step];
                if (!stepValueMap.ContainsKey(row))
                    stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
                var result = new Dictionary<int, object>();
                foreach (var item in stepValueMap[row])
                {
                    result.Add(item.Key, item.Value);
                }
                return result;
            }

        }

        public void SetP2StepState(int stage, int step, int row, bool value)
        {
            CheckP2StepStage(stage);
            _stepP2State[stage][step, row] = value;
        }

        public bool GetP2StepState(int stage, int step, int row)
        {
            CheckP2StepStage(stage);
            return _stepP2State[stage][step, row];
        }


        #endregion

        public ProcessContext(int capacity)
        {
            Capacity = capacity;
            _stepP1State = new bool[StepCont, capacity];

            ValueP1Map = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
            ValueP2Map = new Dictionary<int, Dictionary<int, Dictionary<int, Dictionary<int, object>>>>();
            _stepP2State = new Dictionary<int, bool[,]>();
        }
    }
}