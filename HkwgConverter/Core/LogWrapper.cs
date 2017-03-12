using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HkwgConverter.Core
{
    /// <summary>
    /// Helper class to analyze logging issues on the target machine (with .net 4 client profile only)
    /// Can be removed later
    /// </summary>
    public class LogWrapper 
    {
        Logger nlogger;

        private LogWrapper(Logger nlogLogger)
        {
            this.nlogger = nlogLogger;
        }

        public static LogWrapper GetLogger(Logger nlogLogger)
        {
            return new LogWrapper(nlogLogger);
        }
        /// <summary>
        /// helper function to find out what the program is doing on the target machine
        /// </summary>
        /// <param name="message"></param>
        /// <param name="args"></param>
        public void Error(string message, params object[] args)
        {
            if (Settings.Default.UseConsole)
            {
                Console.WriteLine(message, args);
            }

            nlogger.Error(message, args);
        }

        /// <summary>
        /// helper function to find out what the program is doing on the target machine
        /// </summary>
        /// <param name="message"></param>
        /// <param name="args"></param>
        public void Info(string message, params object[] args)
        {
            if (Settings.Default.UseConsole)
            {
                Console.WriteLine(message, args);
            }

            nlogger.Info(message, args);

        }
    }
}
