using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlueBit.ILF.Reports.ForProjectManagers.Diagnostics
{
    public static class ObjectExtensions
    {
        public static T CheckNotNull<T>(this T @this)
            where T:class
        {
            Contract.Assert(@this != null);
            return @this;
        }
    }
}
