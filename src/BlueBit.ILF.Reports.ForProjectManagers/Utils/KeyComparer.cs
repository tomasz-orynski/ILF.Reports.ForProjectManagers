using System;
using System.Collections.Generic;

namespace BlueBit.ILF.Reports.ForProjectManagers.Utils
{
    public interface IKeyComparer<T> : IEqualityComparer<T>
    {
        T Normalize(T key);
    }

    public abstract class KeyComparer<T> : IKeyComparer<T>
    {
        public int GetHashCode(T obj) => Normalize(obj).GetHashCode();
        public abstract bool Equals(T x, T y);
        public abstract T Normalize(T key);
    }

    public class KeyComparer : KeyComparer<string>
    {
        public override bool Equals(string x, string y) => string.Compare(x, y, true) == 0;
        public override string Normalize(string key) => key.ToLower();
    }
    public class KeyComparerT2 : KeyComparer<(string a, string b)>
    {
        public override bool Equals((string a, string b) x, (string a, string b) y)
            => string.Compare(x.a, y.a, true) == 0
            && string.Compare(x.b, y.b, true) == 0;

        public override (string a, string b) Normalize((string a, string b) key) => (key.a.ToLower(), key.b.ToLower());
    }
    public class KeyComparerT3 : KeyComparer<(string a, string b, string c)>
    {
        public override bool Equals((string a, string b, string c) x, (string a, string b, string c) y)
            => string.Compare(x.a, y.a, true) == 0
            && string.Compare(x.b, y.b, true) == 0
            && string.Compare(x.c, y.c, true) == 0;
        public override (string a, string b, string c) Normalize((string a, string b, string c) key) => (key.a.ToLower(), key.b.ToLower(), key.c.ToLower());
    }
}
