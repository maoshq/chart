using System;
using System.Collections.Generic;
using System.Text;

namespace UITest.Util
{
    public class MyDictionaryComparer : IEqualityComparer<string>
    {
        public bool Equals(string x, string y)
        {
            //throw new NotImplementedException();
            return x != y;
        }

        public int GetHashCode(string obj)
        {
            //throw new NotImplementedException();
            return obj.GetHashCode();
        }
    }
}
