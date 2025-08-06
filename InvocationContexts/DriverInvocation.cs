using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;


namespace Printune
{
    /// <summary>
    /// The DriverInvocation class is for use when installing or uninstalling drivers.
    /// During instantiation, the intent is passed, and this is how it knows whether
    /// to install or uninstall the driver.
    /// This class does not enable printer drivers. That is done when 
    /// adding/installing printers.
    /// </summary>
    public class DriverInvocation : IInvocationContext
    {
        /// <summary>
        /// The name of the driver.
        /// </summary>
        private string _name;
        /// <summary>
        /// Indicates whether the filesystem should be crawled
        /// further recursively in attempt to find nested .inf files.
        /// </summary>
        private bool _recurse;
        /// <summary>
        /// The path to the folder or .inf file.
        /// </summary>
        private string _path;
        /// <summary>
        /// The intent and context of the invocation.
        /// </summary>
        private string _intent;
        private static Dictionary<string, string> _intentStrings = new Dictionary<string, string>();
        /// <summary>
        /// Register the strings that are used to determine the context of the invocation.
        /// </summary>
        public static void Register()
        {
            // Installation Context
            var installHelp = @"printune.exe InstallDriver [-Recurse] {-Path <driver.inf> | -Path <folder\> } { -Name <driver name> } [-LogPath <file.log>]";
            Invocation.RegisterContext("InstallDriver".ToLower(), typeof(DriverInvocation), installHelp);
            _intentStrings.Add("InstallDriver".ToLower(), "installation");

            // Uninstallation Context
            var uninstallHelp = @"printune.exe UninstallDriver { -Driver <PrinterDriverName> | -Path <driver.inf> } [-LogPath <file.log>]";
            Invocation.RegisterContext("UninstallDriver".ToLower(), typeof(DriverInvocation), uninstallHelp);
            _intentStrings.Add("UninstallDriver".ToLower(), "uninstallation");
        }
        /// <summary>
        /// The only constructor, very simple.
        /// </summary>
        /// <param name="Args">The arguments provided from the command line.</param>
        /// <param name="Intent">Arg[0] string parsed to Invocation.Context enum, informs intent of invocation (un/install).</param>
        /// <exception cref="Invocation.ElevatedPrivilegesRequiredException">Thrown when elevated privileges are required; ie: driver installation.</exception>
        /// <exception cref="Invocation.InvalidNameOrPathException">Thrown when an invalid name (ex: driver name) or path (ex: driver path) is provided.</exception>
        public DriverInvocation(string[] Args)
        {
            bool debug = false;

#if DEBUG
            debug = true;
#endif

            if (!Invocation.IsElevated && !debug)
                throw new Invocation.ElevatedPrivilegesRequiredException($"Elevated privilege required for {_intent} of print drivers.");

            _intent = _intentStrings[Args[0].ToLower()];
            _recurse = ParameterParser.GetFlag(Args, "-Recurse");

            // If no path is provided, default to working directory.
            if (!ParameterParser.GetParameterValue(Args, "-Path", out _path))
            {
                _path = Environment.CurrentDirectory;
            }

            if (!ParameterParser.GetParameterValue(Args, "-Name", out _name) && _intent == "installation")
            {
                throw new Invocation.InvalidNameOrPathException("No driver name provided for installation.");
            }

            if (!File.Exists(_path) && !Directory.Exists(_path))
            {
                throw new Invocation.InvalidNameOrPathException("The provided driver path does not exist.");
            }

            // If a file path was provided, but the recurse option
            // was still set for some reason, turn off recurse.
            if (_recurse && File.Exists(_path))
            {
                _recurse = false;
            }
        }

        /// <summary>
        /// Installs or uninstalls driver(s) according to the provided parameters.
        /// </summary>
        /// <returns>0: Success, 1: Error</returns>
        public int Invoke()
        {
            var installationResults = new List<PnpUtil.Result>();
            bool driverEnabled = false;

            // Iterates over all *.inf files found in _path,
            // [un]installing each and keeping track of their output and results.
            foreach (string file in GetDriverInfFiles())
            {
                // Drop the fully-qualified path down to the relative
                // to make logging easier to read.
                var fileRelativePath = file.Replace(_path, "");
                PnpUtil.Result result;

                // Switch the invocation based on provided context.
                if (_intent == "installation")
                    result = PnpUtil.InstallDriver(_path);
                else
                    result = PnpUtil.UninstallDriver(_name ?? _path);

                if (result.Success)
                {
                    Log.Write(fileRelativePath + $": {_intent} succeeded with exit code {result.ExitCode}:");
                    Log.Write($"Enabling print driver '{_name}'...", Indent: 1);
                    try
                    {
                        if (!String.IsNullOrEmpty(_name))
                            driverEnabled |= PrinterDriver.EnablePrinterDriver(_name) != null;
                    }
                    catch
                    {
                        // Don't do anything, just continue.
                    }
                }
                else
                    Log.Write(fileRelativePath + $" {_intent} failed with exit code {result.ExitCode}:", true);

                Log.Write("OUTPUT:");
                Log.Write(result.Output, !result.Success, Indent: 1);

                if (!string.IsNullOrEmpty(result.Error))
                {
                    Log.Write("ERROR:");
                    Log.Write(result.Error, !result.Success, 1);
                }

                installationResults.Add(result);
            }

            // Checks if any pnputil invocations failed.
            if (installationResults.Select(result => result.Success).Contains(false))
            {
                return 1;
            }

            return 0;
        }

        /// <summary>
        /// Finds all *.inf files in _path, recursing if enabled.
        /// </summary>
        /// <returns>Returns a List<string> of *.inf file paths.</returns>
        private List<string> GetDriverInfFiles()
        {
            List<string> fileList = new List<string>();

            if (!_recurse && File.Exists(_path))
            {
                fileList.Add(Path.GetFullPath(_path));
                return fileList;
            }

            SearchOption recursiveSearch = _recurse ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            fileList.AddRange(Directory.GetFiles(_path, "*.inf", recursiveSearch));

            return fileList;
        }
    }
}