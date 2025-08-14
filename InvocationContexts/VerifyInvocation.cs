using System;
using System.Collections.Generic;
using System.IO;

namespace Printune
{
    public class VerifyInvocation : IInvocationContext
    {
        /// <summary>
        /// The name of the printer or driver to verify.
        /// </summary>
        private string _name;
        private string _version;
        private string _intent;
        private static Dictionary<string, string> _intentStrings = new Dictionary<string, string>();

        public VerifyInvocation(string[] Args)
        {
            _intent = _intentStrings[Args[0].ToLower()];

            string paramFilePath = null;
            if (ParameterParser.GetParameterValue("-ParamFile", out paramFilePath))
                ParseParameterFile(paramFilePath);
            else if (File.Exists(Invocation.ParamFile))
                ParseParameterFile(Invocation.ParamFile);
            else
                ParseCommandLine();
        }

        public static void Register()
        {
            var verifyPrinterHelp = @"printune.exe VerifyPrinter -Name <PrinterName> [-LogPath <file.log>]";
            Invocation.RegisterContext("VerifyPrinter".ToLower(), typeof(VerifyInvocation), verifyPrinterHelp);
            _intentStrings.Add("VerifyPrinter".ToLower(), "printer");

            var verifyDriverHelp = @"printune.exe VerifyDriver -Name <DriverName> [-Version <Version>] [-LogPath <file.log>]";
            Invocation.RegisterContext("VerifyDriver".ToLower(), typeof(VerifyInvocation), verifyDriverHelp);
            _intentStrings.Add("VerifyDriver".ToLower(), "driver");
        }

        private void ParseCommandLine()
        {
            if (!ParameterParser.GetParameterValue("-Name", out _name))
                throw new Invocation.InvalidNameOrPathException("No name provided.");

            ParameterParser.GetParameterValue("-Version", out _version);
        }

        private void ParseParameterFile(string parameterFilePath)
        {
            var parameterFile = new ParameterFile(parameterFilePath);

            if (_intent == "printer")
            {
                try
                {
                    _name = (string)parameterFile.GetParameter("PrinterName");
                }
                catch (KeyNotFoundException)
                {
                    throw new Invocation.InvalidNameOrPathException("No \"PrinterName\" found in parameter file.");
                }
                
                return;
            }
            
            try
            {
                _name = (string)parameterFile.GetParameter("Driver");
            }
            catch (KeyNotFoundException)
            {
                throw new Invocation.InvalidNameOrPathException("No \"Name\" found in parameter file.");
            }

            try
            {
                _version = (string)parameterFile.GetParameter("Version");
            }
            catch (KeyNotFoundException)
            {
                // Do nothing, version is optional.
            }
        }
        public int Invoke()
        {
            if (_intent == "printer")
                return VerifyPrinter() ? 0 : 1;
            else
                return VerifyDriver() ? 0 : 1;
        }

        private bool VerifyPrinter()
        {
            using (var printer = Printer.FromExisting(_name))
            {
                Log.Write($"Printer \"{_name}\" {(printer.Exists ? "exists" : "does not exist")}.");
                return printer.Exists;
            }
        }

        private bool VerifyDriver()
        {
            using (var driver = new PrinterDriver(_name, _version))
            {
                Log.Write($"Printer driver \"{_name}\"{(!String.IsNullOrEmpty(_version) ? $", version {_version} " : " ")}{(driver.Exists ? "exists" : "does not exist")}.");
                return driver.Exists;
            }
        }
    }
}