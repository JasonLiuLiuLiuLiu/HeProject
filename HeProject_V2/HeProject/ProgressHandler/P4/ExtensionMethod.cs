using System.Collections.Generic;

namespace HeProject.ProgressHandler.P4
{
    public static class ExtensionMethod
    {
        public static Dictionary<int, object> CheckIntValue(this Dictionary<int, object> dic, int length = 6)
        {
            for (int i = 0; i < 6; i++)
            {
                if (!dic.ContainsKey(i))
                {
                    dic.Add(i, 0);
                }
            }
            return dic;
        }
    }
}
