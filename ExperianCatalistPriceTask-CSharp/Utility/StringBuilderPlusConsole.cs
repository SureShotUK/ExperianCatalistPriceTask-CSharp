using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExperianCatalistPriceTask_CSharp.Utility
{
    public class StringBuilderPlusConsole
    {
        private readonly static StringBuilder LogString = new();
        private readonly static StringBuilder ErrorLogString = new();
        // Writes the inputted string to both a new console line and adds a new line to the StringBuilder string for the emails.
        public static void EmailBodyBuilder(string str)
        {
            Console.WriteLine(str);
            LogString.Append("<p>" + str + "</p>");
        }
        public static void EmailBodyBuilderSBOnly(string str)
        {
            LogString.Append(str);
        }
        public static void ErrorEmailBodyBuilder(string str)
        {
            Console.WriteLine(str);
            ErrorLogString.Append("<p>" + str + "</p>");
        }
        public static void ErrorEmailBodyBuilderSBOnly(string str)
        {
            ErrorLogString.Append(str);
        }
        public static StringBuilder GetLogString()
        {
            return LogString;
        }
        public static StringBuilder GetErrorLogString()
        {
            return ErrorLogString;
        }
    }
}
