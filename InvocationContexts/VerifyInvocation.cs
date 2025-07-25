using System;
using System.Collections.Generic;

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

            if (!ParameterParser.GetParameterValue(Args, "-Name", out _name))
                throw new Invocation.InvalidNameOrPathException("No name provided.");

            ParameterParser.GetParameterValue(Args, "-Version", out _version);
        }

        public static void Register()
        {
            var verifyPrinterHelp = @"printune.exe VerifyPrinter -Name <PrinterName> [-LogPath <path\to\file.log>]";
            Invocation.RegisterContext("VerifyPrinter".ToLower(), typeof(VerifyInvocation), verifyPrinterHelp);
            _intentStrings.Add("VerifyPrinter".ToLower(), "printer");

            var verifyDriverHelp = @"printune.exe VerifyDriver -Name <DriverName> [-Version <Version>] [-LogPath <path\to\file.log>]";
            Invocation.RegisterContext("VerifyDriver".ToLower(), typeof(VerifyInvocation), verifyDriverHelp);
            _intentStrings.Add("VerifyDriver".ToLower(), "driver");
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