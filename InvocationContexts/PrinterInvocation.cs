using System;
using System.Collections.Generic;
using System.IO;
using Printune.Models;

namespace Printune
{
    /// <summary>
    /// The invocation context used when adding/installing or removing/uninstalling printers.
    /// </summary>
    public class PrinterInvocation : IInvocationContext
    {
        /// <summary>
        /// The name of a printer defined in the configuration file.
        /// </summary>
        private string _printerName;
        /// <summary>
        /// The path of the configuration file, default is config.json.
        /// </summary>
        private string _config = Path.Combine(Environment.CurrentDirectory, "config.json");
        // private string _defaultConfig = Path.Combine(Environment.CurrentDirectory, "config.json");
        private string _configContent;
        private string _intent;
        private static Dictionary<string, string> _intentStrings = new Dictionary<string, string>();
        public static void Register()
        {
            // Installation Context
            var installHelp = @"printune.exe InstallPrinter { -PrinterName <PrinterName> } [ -Config <config.json> ] [-LogPath <file.log>]";
            Invocation.RegisterContext("InstallPrinter".ToLower(), typeof(PrinterInvocation), installHelp);
            _intentStrings.Add("InstallPrinter".ToLower(), "installation");

            // Uninstallation Context
            var uninstallHelp = @"printune.exe UninstallPrinter { -PrinterName <PrinterName> } [-LogPath <file.log>]";
            Invocation.RegisterContext("UninstallPrinter".ToLower(), typeof(PrinterInvocation), uninstallHelp);
            _intentStrings.Add("UninstallPrinter".ToLower(), "uninstallation");
        }
        public PrinterInvocation(string[] Args)
        {
            _intent = _intentStrings[Args[0].ToLower()];

            string paramFile = string.Empty;
            if (ParameterParser.GetParameterValue("-ParamFile", out paramFile))
                ParseParameterFile(paramFile);
            else if (File.Exists(Invocation.ParamFile))
                ParseParameterFile(Invocation.ParamFile);                
            else
                ParseCommandLine();

            if (_intent == "installation")
            {
                // Reads the config. ConfigReader parses the path
                // and determines whether it's a local/SMB file
                // or if it's an HTTP/S URI.
                _configContent = ConfigReader.Read(_config);
                Log.Write($"Content read from {_config} configuration file.");
            }
        }

        private void ParseCommandLine()
        {
            if (!ParameterParser.GetParameterValue("-PrinterName", out _printerName))
            {
                if (!ParameterParser.GetParameterValue("-Name", out _printerName))
                    throw new Invocation.MissingArgumentException("Invalid invocation: 'PrinterName' paramer is required.");
            }

            // If the Config parameter is provided a value, use the default.
            if (!ParameterParser.GetParameterValue("-Config", out _config) && _intent == "installation")
            {
                // If the default does not exist and no other has been specified, throw an error.
                if (!File.Exists(_config))
                {
                    throw new Invocation.MissingArgumentException(
                        "Invalid invocation: A valid URI must be passed with the 'Config' parameter"
                        + " or there must be a 'config.json' file in the current working directory."
                        );
                }
            }
        }

        private void ParseParameterFile(string paramFilePath)
        {
            var paramFile = new ParameterFile(paramFilePath);
            try
            {
                _printerName = (string)paramFile.GetParameter("PrinterName");
            }
            catch (KeyNotFoundException)
            {
                throw new Invocation.InvalidNameOrPathException("The \"Name\" parameter is required in the parameter file.");
            }

            try
            {
                _config = (string)paramFile.GetParameter("Config");
            }
            catch (KeyNotFoundException)
            {
                // If the default does not exist and no other has been specified, throw an error.
                if (!File.Exists(_config))
                {
                    throw new Invocation.MissingArgumentException(
                        "Invalid invocation: A valid URI must be passed with the 'Config' parameter"
                        + " or there must be a 'config.json' file in the current working directory."
                    );
                }
            }
        }
        /// <summary>
        /// Invokes the necessary methods to install or uninstall the printer.
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Invocation.ConfigurationFileException"></exception>
        public int Invoke()
        {
            if (_intent == "installation")
            {
                if (string.IsNullOrEmpty(_configContent))
                    throw new Invocation.ConfigurationFileException("Configuration file content empty.");

                // Try to instantiate a Printer from the JSON. This does not create
                // the Windows print queue, only an object representing one.
                var printer = Printer.FromConfig(_printerName, _configContent);
                Log.Write($"Printer \"{_printerName}\" successfully hydrated from configuration file.");

                var driver = new PrinterDriver(printer.DriverName);
                if (!driver.Exists)
                    throw new Invocation.ConfigurationFileException($"Printer driver \"{printer.DriverName}\" not installed and/or enabled.");

                if (!PrintProcessor.Exists(printer.PrintProcessor))
                    throw new Invocation.ConfigurationFileException($"Print Processes \"{printer.PrintProcessor}\" does not exist on system.");

                // Creates Windows the print queue.
                    printer.Commit();
                Log.Write($"Printer successfully added.");

                if (!String.IsNullOrEmpty(printer.PreferenceFile))
                    new PrinterPreference(printer.PreferenceFile).Apply(_printerName);
            }
            else
            {
                // Attempts to instantiate a Printer object.
                var printer = Printer.FromExisting(_printerName);

                if (printer == null)
                    Log.Write($"Printer objected name \"{_printerName}\" could not be found for removal.", true);
                else
                {
                    Log.Write("Printer found in system.");
                    // Remove Windows print queue.
                    printer.Remove();
                    Log.Write("Printer successfully removed.");
                }
            }
            return 0;
        }
    }
}