using NLog;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ILF.Reports.ForProjectManagers.Diagnostics
{
    public static class LoggerExtensions
    {
        public static void EntryCall(Logger logger, Action action, string name)
        {
            try
            {
                logger.Trace(">>" + name);
                action();
                logger.Trace("<<" + name);
            }
            catch (Exception e)
            {
                logger.Error(e, "!!" + name);
            }
        }
        public static void EntryCall(Logger logger, Action action) => EntryCall(logger, action, action.Method.Name);

        public static void OnEntryCall(this Logger @this, Action action, [CallerMemberName]string name = null)
        {
            Contract.Assert(@this != null);
            Contract.Assert(action != null);
            Contract.Assert(!string.IsNullOrEmpty(name));
            EntryCall(@this, action, name);
        }

        public static T OnEntryCall<T>(this Logger @this, Func<T> action, [CallerMemberName]string name = null)
        {
            Contract.Assert(@this != null);
            Contract.Assert(action != null);
            Contract.Assert(!string.IsNullOrEmpty(name));
            var result = default(T);
            EntryCall(@this, () => result = action(), name);
            return result;
        }
    }
}
