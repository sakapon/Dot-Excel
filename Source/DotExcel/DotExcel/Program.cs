using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using Keiho.Apps.DotExcel.Consoles;

namespace Keiho.Apps.DotExcel
{
    static class Program
    {
        static int Main(string[] args)
        {
            try
            {
                var args2 = Commands.CorrectArgs(args);
                CommandUtility.Execute(args2);
                return 0;
            }
            catch (CommandArgumentsException)
            {
                CommandUtility.ShowHelp();
                return 0;
            }
            catch (WarningException ex)
            {
                Console.WriteLine(ex.Message);
                return 10;
            }
            catch (AggregateException ex)
            {
                ex.InnerExceptions.ForEach(iex =>
                {
                    if (iex is WarningException)
                    {
                        Console.WriteLine(iex.Message);
                    }
                    else
                    {
                        Console.WriteLine("予期しないエラーが発生しました。");
                        Console.WriteLine(iex);
                    }
                    Console.WriteLine();
                })
                .Execute();

                return ex.InnerExceptions.All(iex => iex is WarningException) ? 10 : 100;
            }
            catch (Exception ex)
            {
                Console.WriteLine("予期しないエラーが発生しました。");
                Console.WriteLine(ex);
                return 100;
            }
        }
    }
}
