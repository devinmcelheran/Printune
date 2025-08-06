using System;
using System.Collections.Generic;

namespace Printune
{
    /// <summary>
    /// The fallback invocation context. If an invocation is invalid,
    /// an exception is thrown, which should invoke the help context.
    /// </summary>
    public class HelpInvocation : IInvocationContext
    {
        /// <summary>
        /// Error message and/or error help message is displayed when true.
        /// </summary>
        private bool _errorHelp = false;
        /// <summary>
        /// Error and/or error help message content.
        /// </summary>
        private string _message;
        /// <summary>
        /// The integer that Program.cs->Main() should exit with.
        /// </summary>
        private int _exitCode = 0;
        /// <summary>
        /// Standard help message (_helpString) is displayed when true.
        /// </summary>
        private bool _includeHelp = true;
        /// <summary>
        /// JSON template/schema for printer installation is displayed when true.
        /// </summary>
        private bool _includeTemplate = false;
        /// <summary>
        /// JSON template/schema content.
        /// </summary>
        private string _templateString =
    @"[
    {
        ""PrinterName"": ""Hallway Printer"",
        ""Version"": ""1.0"",
        ""DriverName"": ""Zerocks Universal PCL"",
        ""DataType"": ""RAW"",
        ""PrintProcessor"": ""winprint"",
        ""Location"": ""In the hallway."",
        ""PrinterPort"": {
        ""PrinterHostAddress"": ""1.2.3.4"",
        ""PortNumber"": 9100,
        ""SNMP"": 1,
        ""SNMPCommunity"": ""public""
        }
    },
    {
        ""PrinterName"": ""Other Hallway Printer"",
        ""Version"": ""1.0"",
        ""DriverName"": ""Zerocks Universal PCL"",
        ""DataType"": ""RAW"",
        ""PrintProcessor"": ""winprint"",
        ""Location"": ""In the hallway."",
        ""PrinterPort"": {
        ""PrinterHostAddress"": ""1.2.3.5"",
        ""PortNumber"": 9100,
        ""SNMP"": 1,
        ""SNMPCommunity"": ""public""
        }
    }
]";

        public static void Register()
        {
            var help = @"printune.exe Help [-Template]";
            Invocation.RegisterContext("Help".ToLower(), typeof(HelpInvocation), help);
        }

        public HelpInvocation() { }
        /// <summary>
        /// Default invocation.
        /// </summary>
        /// <param name="Args">Commandline arguments.</param>
        public HelpInvocation(string[] Args)
        {
            // If the template is specifically requested, we don't need
            // to pollute the screen with a bunch more text.
            _includeTemplate = ParameterParser.GetFlag(Args, "-Template");
            if (_includeTemplate)
                _includeHelp = false;
        }
        /// <summary>
        /// I don't think this is in use anymore.
        /// </summary>
        /// <param name="Message">Message to be shown.</param>
        /// <param name="ExitCode">The int Program.cs->Main() exits with.</param>
        /// <param name="ShowTemplate">Shows JSON template when true.</param>
        public HelpInvocation(string Message, int ExitCode = 0, bool ShowTemplate = false)
        {
            _message = Message;
            _includeTemplate = ShowTemplate;
        }

        /// <summary>
        /// When falling back from an invocation error, this constructor is used.
        /// The Exception parameter is used to determine what and how the help
        /// message ends up being displayed.
        /// </summary>
        /// <param name="Ex"></param>
        public HelpInvocation(Exception Ex)
        {
            _errorHelp = true;
            _exitCode = 1;

            _message = Ex.Message;

            // When running in debug, print the whole error message,
            // stack trace and all, for convenience.
#if DEBUG
            _message = Ex.ToString();
#endif

            switch (Ex)
            {
                case Invocation.ConfigurationFileException ex:
                    _includeTemplate = true;
                    _includeHelp = false;
                    break;
                case Invocation.InvalidNameOrPathException ex:
                    _includeHelp = false;
                    break;
                case Invocation.MissingArgumentException ex:
                    // Show error message with command syntax
                    break;
                case Invocation.ElevatedPrivilegesRequiredException ex:
                    // Show error message
                    _includeHelp = false;
                    break;
                case Newtonsoft.Json.JsonException ex:
                    _includeTemplate = true;
                    _includeHelp = false;
                    break;
                default:
                    _message = $"An unhandled error occured:\n\n{Ex}";
                    break;
            }
        }

        /// <summary>
        /// Invocation of this context simply prints help message
        /// and errors for the user.
        /// </summary>
        /// <returns></returns>
        public int Invoke()
        {
            if (_message != null)
                Log.Write(_message, _errorHelp);

            if (_includeTemplate)
                Log.Write(_templateString);

            if (_includeHelp)
                Log.Write(Invocation.HelpText + "\nprintune.exe Help [-Template]");

            return _exitCode;
        }
    }
}