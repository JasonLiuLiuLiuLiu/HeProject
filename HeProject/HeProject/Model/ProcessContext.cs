using System;
using System.Collections.Generic;
using System.Threading;

namespace HeProject.Model
{
    public class ProcessContext
    {
        private readonly bool[,] _stepSourceState;

        //private readonly bool[,] _stepSourceP2State;
        //private readonly bool[,] _stepSourceP3State;
        private readonly bool[,] _stepP1State;
        private readonly bool[,] _stepP2State;
        private readonly bool[,] _stepP3State;
        private readonly bool[,] _stepP4State;
        private readonly bool[,] _stepP5State;

        public readonly Dictionary<int, Dictionary<int, Dictionary<int, object>>> _valueSourceMap;

        //private readonly Dictionary<int, Dictionary<int, Dictionary<int, object>>> _valueSourceP2Map;
        //private readonly Dictionary<int, Dictionary<int, Dictionary<int, object>>> _valueSourceP3Map;
        private readonly Dictionary<int, Dictionary<int, Dictionary<int, object>>> _valueP1Map;
        private readonly Dictionary<int, Dictionary<int, Dictionary<int, object>>> _valueP2Map;
        private readonly Dictionary<int, Dictionary<int, Dictionary<int, object>>> _valueP3Map;
        private readonly Dictionary<int, Dictionary<int, Dictionary<int, object>>> _valueP4Map;
        private readonly Dictionary<int, Dictionary<int, Dictionary<int, object>>> _valueP5Map;
        public readonly int Capacity;
        private const int StepCont = 9;
        private readonly object _lockSourceP1 = new object();

        //private readonly object _lockSourceP2 = new object();
        //private readonly object _lockSourceP3 = new object();
        private readonly object _lockP1 = new object();
        private readonly object _lockP2 = new object();
        private readonly object _lockP3 = new object();
        private readonly object _lockP4 = new object();
        private readonly object _lockP5 = new object();

        #region Source

        public void SetSourceValue(int step, int row, int column, object value)
        {
            try
            {
                lock (_lockSourceP1)
                {
                    if (!_valueSourceMap.ContainsKey(step))
                        _valueSourceMap.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = _valueSourceMap[step];
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

        public T GetSourceValue<T>(int step, int row, int column)
        {
            try
            {
                int loop = 0;
                while (!_stepSourceState[step - 1, row])
                {
                    if (loop > 100)
                        return default(T);
                    Thread.Sleep(50);
                }
                lock (_lockSourceP1)
                {
                    if (!_valueSourceMap.ContainsKey(step))
                        _valueSourceMap.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = _valueSourceMap[step];
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

        public Dictionary<int, object> GetSourceRowResult(int step, int row)
        {
            int loop = 0;
            while (!_stepSourceState[step - 1, row])
            {
                if (loop > 100)
                    return new Dictionary<int, object>(StepLength.SourceLength);
                Thread.Sleep(50);
            }
            lock (_lockSourceP1)
            {
                if (!_valueSourceMap.ContainsKey(step))
                    _valueSourceMap.Add(step, new Dictionary<int, Dictionary<int, object>>(StepLength.SourceLength));
            }

            var stepValueMap = _valueSourceMap[step];
            if (!stepValueMap.ContainsKey(row))
                stepValueMap.Add(row, new Dictionary<int, object>(StepLength.SourceLength));
            return stepValueMap[row];
        }

        public void SetSourceStepState(int step, int row, bool value)
        {
            _stepSourceState[step - 1, row] = value;
        }

        #endregion Source

        //#region SourceSourceP2

        //public void SetSourceP2Value(int step, int row, int column, object value)
        //{
        //    try
        //    {
        //        lock (_lockSourceP2)
        //        {
        //            if (!_valueSourceP2Map.ContainsKey(step))
        //                _valueSourceP2Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
        //            var stepValueMap = _valueSourceP2Map[step];
        //            if (!stepValueMap.ContainsKey(row))
        //                stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
        //            var rowValueMap = stepValueMap[row];
        //            if (!rowValueMap.ContainsKey(column))
        //                rowValueMap.Add(column, null);

        //            rowValueMap[column] = value;
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine($"step:{step},row:{row},column:{column},value:{value},exception:{e}");
        //        throw;
        //    }
        //}

        //public T GetSourceP2Value<T>(int step, int row, int column)
        //{
        //    try
        //    {
        //        lock (_lockSourceP2)
        //        {
        //            int loop = 0;
        //            while (!_stepSourceP2State[step - 1, row])
        //            {
        //                if (loop > 100)
        //                    return default(T);
        //                Thread.Sleep(50);
        //            }
        //            if (!_valueSourceP2Map.ContainsKey(step))
        //                _valueSourceP2Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
        //            var stepValueMap = _valueSourceP2Map[step];
        //            if (!stepValueMap.ContainsKey(row))
        //                stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
        //            var rowValueMap = stepValueMap[row];
        //            if (!rowValueMap.ContainsKey(column))
        //                rowValueMap.Add(column, default(T));
        //            return (T)rowValueMap[column];
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine($"step:{step},row:{row},column:{column},exception:{e}");
        //        throw;
        //    }
        //}

        //public void SetSourceP2StepState(int step, int row, bool value)
        //{
        //    _stepSourceP2State[step - 1, row] = value;
        //}

        //#endregion SourceSourceP2

        //#region SourceSourceP3

        //public void SetSourceP3Value(int step, int row, int column, object value)
        //{
        //    try
        //    {
        //        lock (_lockSourceP3)
        //        {
        //            if (!_valueSourceP3Map.ContainsKey(step))
        //                _valueSourceP3Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
        //            var stepValueMap = _valueSourceP3Map[step];
        //            if (!stepValueMap.ContainsKey(row))
        //                stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
        //            var rowValueMap = stepValueMap[row];
        //            if (!rowValueMap.ContainsKey(column))
        //                rowValueMap.Add(column, null);

        //            rowValueMap[column] = value;
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine($"step:{step},row:{row},column:{column},value:{value},exception:{e}");
        //        throw;
        //    }
        //}

        //public T GetSourceP3Value<T>(int step, int row, int column)
        //{
        //    try
        //    {
        //        lock (_lockSourceP3)
        //        {
        //            int loop = 0;
        //            while (!_stepSourceP3State[step - 1, row])
        //            {
        //                if (loop > 100)
        //                    return default(T);
        //                Thread.Sleep(50);
        //            }
        //            if (!_valueSourceP3Map.ContainsKey(step))
        //                _valueSourceP3Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
        //            var stepValueMap = _valueSourceP3Map[step];
        //            if (!stepValueMap.ContainsKey(row))
        //                stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
        //            var rowValueMap = stepValueMap[row];
        //            if (!rowValueMap.ContainsKey(column))
        //                rowValueMap.Add(column, default(T));
        //            return (T)rowValueMap[column];
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine($"step:{step},row:{row},column:{column},exception:{e}");
        //        throw;
        //    }
        //}

        //public void SetSourceP3StepState(int step, int row, bool value)
        //{
        //    _stepSourceP3State[step - 1, row] = value;
        //}

        //#endregion SourceSourceP3

        #region P1

        public void SetP1Value(int step, int row, int column, object value)
        {
            try
            {
                lock (_lockP1)
                {
                    if (!_valueP1Map.ContainsKey(step))
                        _valueP1Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = _valueP1Map[step];
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
                while (!_stepP1State[step - 1, row])
                {
                    if (loop > 100)
                        return default(T);
                    Thread.Sleep(50);
                }
                lock (_lockP1)
                {
                    if (!_valueP1Map.ContainsKey(step))
                        _valueP1Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = _valueP1Map[step];
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
            while (!_stepP1State[step - 1, row])
            {
                if (loop > 100)
                    return new Dictionary<int, object>(Capacity);
                Thread.Sleep(50);
            }
            lock (_lockP1)
            {
                if (!_valueP1Map.ContainsKey(step))
                    _valueP1Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
            }

            var stepValueMap = _valueP1Map[step];
            if (!stepValueMap.ContainsKey(row))
                stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
            return stepValueMap[row];
        }

        public void SetP1StepState(int step, int row, bool value)
        {
            _stepP1State[step - 1, row] = value;
        }

        #endregion P1

        #region P2

        public void SetP2Value(int step, int row, int column, object value)
        {
            try
            {
                lock (_lockP2)
                {
                    if (!_valueP2Map.ContainsKey(step))
                        _valueP2Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = _valueP2Map[step];
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

        public T GetP2Value<T>(int step, int row, int column)
        {
            try
            {
                int loop = 0;
                while (!_stepP2State[step - 1, row])
                {
                    if (loop > 100)
                        return default(T);
                    Thread.Sleep(1000);
                }
                lock (_lockP2)
                {
                    if (!_valueP2Map.ContainsKey(step))
                        _valueP2Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = _valueP2Map[step];
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

        public Dictionary<int, object> GetP2RowResult(int step, int row)
        {
            int loop = 0;
            while (!_stepP2State[step - 1, row])
            {
                if (loop > 100)
                    return new Dictionary<int, object>(Capacity);
                Thread.Sleep(1000);
            }
            lock (_lockP2)
            {
                if (!_valueP2Map.ContainsKey(step))
                    _valueP2Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
            }
            var stepValueMap = _valueP2Map[step];
            if (!stepValueMap.ContainsKey(row))
                stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
            return stepValueMap[row];
        }

        public void SetP2StepState(int step, int row, bool value)
        {
            _stepP2State[step - 1, row] = value;
        }

        #endregion P2

        #region P3

        public void SetP3Value(int step, int row, int column, object value)
        {
            try
            {
                lock (_lockP3)
                {
                    if (!_valueP3Map.ContainsKey(step))
                        _valueP3Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = _valueP3Map[step];
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

        public T GetP3Value<T>(int step, int row, int column)
        {
            try
            {
                int loop = 0;
                while (!_stepP3State[step - 1, row])
                {
                    if (loop > 1000)
                        return default(T);
                    Thread.Sleep(1000);
                }
                lock (_lockP3)
                {
                    if (!_valueP3Map.ContainsKey(step))
                        _valueP3Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = _valueP3Map[step];
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

        public Dictionary<int, object> GetP3RowResult(int step, int row)
        {
            int loop = 0;
            while (!_stepP3State[step - 1, row])
            {
                if (loop > 1000)
                    return new Dictionary<int, object>(Capacity);
                Thread.Sleep(1000);
            }
            lock (_lockP3)
            {
                if (!_valueP3Map.ContainsKey(step))
                    _valueP3Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                var stepValueMap = _valueP3Map[step];
                if (!stepValueMap.ContainsKey(row))
                    stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
                return stepValueMap[row];
            }
        }

        public void SetP3StepState(int step, int row, bool value)
        {
            _stepP3State[step - 1, row] = value;
        }

        #endregion P3

        #region P4

        public void SetP4Value(int step, int row, int column, object value)
        {
            try
            {
                lock (_lockP4)
                {
                    if (!_valueP4Map.ContainsKey(step))
                        _valueP4Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = _valueP4Map[step];
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

        public T GetP4Value<T>(int step, int row, int column)
        {
            try
            {
                int loop = 0;
                while (!_stepP4State[step - 1, row])
                {
                    if (loop > 100)
                        return default(T);
                    Thread.Sleep(1000);
                }
                lock (_lockP4)
                {
                    if (!_valueP4Map.ContainsKey(step))
                        _valueP4Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = _valueP4Map[step];
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

        public Dictionary<int, object> GetP4RowResult(int step, int row)
        {
            int loop = 0;
            while (!_stepP4State[step - 1, row])
            {
                if (loop > 100)
                    return new Dictionary<int, object>(Capacity);
                Thread.Sleep(1000);
            }
            lock (_lockP4)
            {
                if (!_valueP4Map.ContainsKey(step))
                    _valueP4Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                var stepValueMap = _valueP4Map[step];
                if (!stepValueMap.ContainsKey(row))
                    stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
                return stepValueMap[row];
            }
        }

        public void SetP4StepState(int step, int row, bool value)
        {
            _stepP4State[step - 1, row] = value;
        }

        #endregion P4

        #region P5

        public void SetP5Value(int step, int row, int column, object value)
        {
            try
            {
                lock (_lockP5)
                {
                    if (!_valueP5Map.ContainsKey(step))
                        _valueP5Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = _valueP5Map[step];
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

        public T GetP5Value<T>(int step, int row, int column)
        {
            try
            {
                int loop = 0;
                while (!_stepP5State[step - 1, row])
                {
                    if (loop > 100)
                        return default(T);
                    Thread.Sleep(1000);
                }
                lock (_lockP5)
                {
                    if (!_valueP5Map.ContainsKey(step))
                        _valueP5Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                    var stepValueMap = _valueP5Map[step];
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

        public Dictionary<int, object> GetP5RowResult(int step, int row)
        {
            int loop = 0;
            while (!_stepP5State[step - 1, row])
            {
                if (loop > 100)
                    return new Dictionary<int, object>(Capacity);
                Thread.Sleep(1000);
            }
            lock (_lockP5)
            {
                if (!_valueP5Map.ContainsKey(step))
                    _valueP5Map.Add(step, new Dictionary<int, Dictionary<int, object>>(StepCont));
                var stepValueMap = _valueP5Map[step];
                if (!stepValueMap.ContainsKey(row))
                    stepValueMap.Add(row, new Dictionary<int, object>(Capacity));
                return stepValueMap[row];
            }
        }

        public void SetP5StepState(int step, int row, bool value)
        {
            _stepP5State[step - 1, row] = value;
        }

        #endregion P5

        public ProcessContext(int capacity)
        {
            Capacity = capacity;
            _stepSourceState = new bool[21, capacity];
            //_stepSourceP2State = new bool[StepCont, capacity];
            //_stepSourceP3State = new bool[StepCont, capacity];
            _stepP1State = new bool[StepCont, capacity];
            _stepP2State = new bool[StepCont, capacity];
            _stepP3State = new bool[StepCont, capacity];
            _stepP4State = new bool[StepCont, capacity];
            _stepP5State = new bool[StepCont, capacity];
            _valueSourceMap = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
            //_valueSourceP2Map = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
            //_valueSourceP3Map = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
            _valueP1Map = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
            _valueP2Map = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
            _valueP3Map = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
            _valueP4Map = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
            _valueP5Map = new Dictionary<int, Dictionary<int, Dictionary<int, object>>>();
        }
    }
}