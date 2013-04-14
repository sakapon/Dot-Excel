using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Keiho.Apps.DotExcel.Consoles
{
    [AttributeUsage(AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
    public sealed class CommandsClassAttribute : Attribute
    {
    }

    [AttributeUsage(AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
    [DebuggerDisplay(@"\{{Message}\}")]
    public sealed class CommandUsageAttribute : Attribute
    {
        public string Message { get; private set; }

        public CommandUsageAttribute(string message)
        {
            Message = message;
        }
    }
}
