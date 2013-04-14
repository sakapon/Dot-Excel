using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Keiho.Apps.DotExcel.Consoles
{
    public static class CommandUtility
    {
        public static void Execute(string[] args)
        {
            if (args == null) throw new ArgumentNullException("args");

            var commandsType = GetMarkedType<CommandsClassAttribute>();
            if (commandsType == null) throw new InvalidOperationException("CommandsClassAttribute でマークされたクラスが存在しません。");

            var commandArgs = new CommandArguments(args);
            if (commandArgs.CommandName == null) throw new CommandArgumentsException();

            var command = commandsType.GetMethods()
                .SingleOrDefault(m => string.Equals(m.Name, commandArgs.CommandName, StringComparison.InvariantCultureIgnoreCase));
            if (command == null) throw new CommandArgumentsException();

            var parameters = command.GetParameters()
                .Select((p, i) =>
                {
                    if (i == 0)
                    {
                        return p.ParameterType == typeof(string[]) ? (object)commandArgs.DefaultArgs : commandArgs.DefaultArgs.FirstOrDefault();
                    }
                    else
                    {
                        var key = p.Name.ToLowerInvariant();
                        return commandArgs.OptionalArgs.ContainsKey(key) ? commandArgs.OptionalArgs[key] : p.DefaultValue;
                    }
                })
                .ToArray();

            try
            {
                command.Invoke(null, parameters);
            }
            catch (TargetInvocationException ex)
            {
                throw ex.InnerException;
            }
        }

        public static void ShowHelp()
        {
            var commandsType = GetMarkedType<CommandsClassAttribute>();
            if (commandsType == null) throw new InvalidOperationException("CommandsClassAttribute でマークされたクラスが存在しません。");

            var usage = commandsType.GetCustomAttributes(false).OfType<CommandUsageAttribute>().SingleOrDefault();
            if (usage != null)
            {
                Console.WriteLine(usage.Message);
            }
        }

        public static Type GetMarkedType<TAttribute>() where TAttribute : Attribute
        {
            return Assembly.GetEntryAssembly().GetTypes()
                .FirstOrDefault(t => t.GetCustomAttributes(typeof(TAttribute), false).Length > 0);
        }

        public static string GetEntryAssemblyPath()
        {
            return Assembly.GetEntryAssembly().Location;
        }

        public static WarningException ToWarning(this Exception ex)
        {
            return new WarningException(ex.Message, ex);
        }
    }
}
