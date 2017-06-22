using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ILF.Reports.ForProjectManagers.Diagnostics
{
    public static class StringExtensions
    {
        public static string NullTrim(this string @this)
        {
            return string.IsNullOrWhiteSpace(@this)
                ? string.Empty
                : @this.Trim();
        }
    }
}
