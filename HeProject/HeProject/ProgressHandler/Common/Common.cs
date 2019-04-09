namespace HeProject.ProgressHandler.Common
{
    public static class SCommon
    {
        public static int CalculateSameLength(this int[] inputs)
        {
            var result = 0;
            if (inputs.Length < 4)
                return result;
            if (inputs[0].IsMin())
            {
                if (inputs[1].IsMin() && inputs[2].IsMin() && inputs[3].IsMin())
                    return 3;
                if (inputs[1].IsMin() && inputs[2].IsMin())
                    return 2;
                if (inputs[1].IsMin())
                    return 1;
            }
            else
            {
                if (inputs[1].IsBig() && inputs[2].IsBig() && inputs[3].IsBig())
                    return 3;
                if (inputs[1].IsBig() && inputs[2].IsBig())
                    return 2;
                if (inputs[1].IsBig())
                    return 1;
            }
            return result;
        }

        private static bool IsMin(this int value)
        {
            return value <= 1;
        }

        private static bool IsBig(this int value)
        {
            return !value.IsMin();
        }
    }
}
