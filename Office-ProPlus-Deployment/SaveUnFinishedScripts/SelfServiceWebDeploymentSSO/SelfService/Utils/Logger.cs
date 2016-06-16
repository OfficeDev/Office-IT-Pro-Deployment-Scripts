#region

using System;
using System.Diagnostics;

#endregion

namespace SelfService.Utils
{
    /// <summary>
    ///     Trace based logger
    /// </summary>
    public class Logger
    {
        /// <summary>
        ///     Log errors and exceptions.
        /// </summary>
        /// <param name="message">Formatted message.</param>
        /// <param name="args">Message arguments.</param>
        public static void Error(string message, params object[] args)
        {
            Trace.TraceError(message, args);
            Console.WriteLine(message, args);
        }

        /// <summary>
        ///     Log warnings.
        /// </summary>
        /// <param name="message">Formatted message.</param>
        /// <param name="args">Message arguments.</param>
        public static void Warning(string message, params object[] args)
        {
            Trace.TraceWarning(message, args);
            Console.WriteLine(message, args);
        }

        /// <summary>
        ///     Log information.
        /// </summary>
        /// <param name="message">Formatted message.</param>
        /// <param name="args">Message arguments.</param>
        public static void Info(string message, params object[] args)
        {
            Trace.TraceInformation(message, args);
            Console.WriteLine(message, args);
        }
    }
}