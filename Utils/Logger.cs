using System;
using System.IO;

namespace Printune
{
    /// <summary>
    /// A simple logging class.
    /// </summary>
    public static class Log {
        private static bool _initialized = false;
        private static string _logPath = System.IO.Path.Combine(Environment.CurrentDirectory, "printune.log");

        /// <summary>
        /// Initializes logger with a file path.
        /// </summary>
        /// <param name="LogPath">The path of the log file.</param>
        public static void Initialize(string LogPath)
        {
            if (LogPath == null)
                throw new NullReferenceException("A null value was provided instead of a valid path.");
                
            _logPath = LogPath ?? _logPath;
            _initialized = true;
        }

        /// <summary>
        /// Writes to console and log file if one has been initialized.
        /// </summary>
        /// <param name="Message">The message to be logged.</param>
        public static void Write(string Message, bool IsError = false, int Indent = 0)
        {
            Message = FormatEntry(Message, Indent);
        
            if (IsError) Console.ForegroundColor = ConsoleColor.Red;
            Console.Write(Message);
            Console.ResetColor();
        
            if (!_initialized) return;

            File.WriteAllText(_logPath, Message);
        }

        private static string FormatEntry(string Entry, int Indent)
        {
            var indentString = "\n" + new string(' ', 4 * Indent);
            return Entry.Trim().Replace("\n", indentString) + "\n\n";
        }
    }
}