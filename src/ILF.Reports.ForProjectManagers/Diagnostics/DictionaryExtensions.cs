using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ILF.Reports.ForProjectManagers.Diagnostics
{
    public static class DictionaryExtensions
    {
        public static void IfExistsValue<TKey,TValue>(this IReadOnlyDictionary<TKey,TValue> @this, TKey key, Action<TValue> action)
        {
            Contract.Assert(@this != null);
            Contract.Assert(action != null);

            TValue value;
            if (@this.TryGetValue(key, out value))
                action(value);
        }
    }
}
