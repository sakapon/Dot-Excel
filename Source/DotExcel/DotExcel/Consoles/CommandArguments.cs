using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Text.RegularExpressions;

namespace Keiho.Apps.DotExcel.Consoles
{
    public class CommandArguments
    {
        public string CommandName { get; private set; }
        public string[] DefaultArgs { get; private set; }
        public Dictionary<string, string> OptionalArgs { get; private set; }

        public CommandArguments(string[] args)
        {
            if (args == null) throw new ArgumentNullException("args");

            CommandName = args.FirstOrDefault();
            DefaultArgs = args.Skip(1).TakeWhile(s => !Regex.IsMatch(s, "^[/-].*$")).ToArray();
            OptionalArgs = args.Skip(1 + DefaultArgs.Length)
                .Select(s =>
                {
                    var m = Regex.Match(s, "^[/-](.+?):(.*)$");
                    if (!m.Success) throw new CommandArgumentsException();

                    return new { Name = m.Groups[1].Value, Value = m.Groups[2].Value };
                })
                .ToDictionary(x => x.Name.ToLowerInvariant(), x => x.Value);
        }
    }

    [Serializable]
    public class CommandArgumentsException : Exception
    {
        public CommandArgumentsException() { }
        public CommandArgumentsException(string message) : base(message) { }
        public CommandArgumentsException(string message, Exception inner) : base(message, inner) { }
        protected CommandArgumentsException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
