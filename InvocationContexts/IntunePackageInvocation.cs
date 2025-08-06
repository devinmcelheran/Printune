using System;
using System.Collections.Generic;
using System.IO;

namespace Printune
{
    /// <summary>
    /// This invocation context is for preparing and, optionally, packaging printers and drivers for Intune.
    /// </summary>
    public class IntunePackageInvocation : IInvocationContext
    {
        /// <summary>
        /// The name or path of the driver when packagine a driver.
        /// </summary>
        private string _driverName;
        /// <summary>
        /// The name of the printer to be packaged.
        /// </summary>
        private string _printerName;
        /// <summary>
        /// Is used to fork toward exporting a system driver or packaging one from the filesystem.
        /// </summary>
        private bool _exportDriverFromSystem = false;
        /// <summary>
        /// The folder within which resources will be gathered for packaging.
        /// </summary>
        private string _outputPath;
        /// <summary>
        /// A *.intunewin package is created when this is true and either _intuneWinUtilPath is valid or an IntuneWinAppUtil.exe instance can be found somewhere along %PATH%.
        /// </summary>
        private bool _createPackage;
        /// <summary>
        /// The path to the IntuneWinAppUtil.exe to use, as provided by the user.
        /// </summary>
        private string _intuneWinUtilPath = null;
        /// <summary>
        /// This conflates the driver name and printer name to make for simpler reference later on.
        /// </summary>
        private string _packageName
        {
            get
            {
                if (_driverName != null) return Path.GetFileName(_driverName).Replace(".inf", "");
                if (_printerName != null) return _printerName;

                throw new Invocation.MissingArgumentException("Invalid invocation: A valid Driver or PrinterName value must be provided.");
            }
        }
        /// <summary>
        /// The context of the invocation (printer/driver).
        /// </summary>
        private string _intent;
        private static Dictionary<string, string> _intentStrings = new Dictionary<string, string>();
        public static void Register()
        {
            // Driver Context
            var driverHelp = @"printune.exe PackageDriver { -Driver <PrinterDriverName> | -Path <driver.inf> } [ -Output <destination\> ] [ -IntuneWinUtil <path\intunewinutil.exe> ] [-LogPath <path\to\file.log>]";
            Invocation.RegisterContext("PackageDriver".ToLower(), typeof(IntunePackageInvocation), driverHelp);
            _intentStrings.Add("PackageDriver".ToLower(), "driver");
            
            // Printer Context
            var printerHelp = @"printune.exe PackagePrinter { -PrinterName <PrinterName>} [ -Output <destination\> ] [ -IntuneWinUtil <path\intunewinutil.exe> ] [-LogPath <file.log>]";
            Invocation.RegisterContext("PackagePrinter".ToLower(), typeof(IntunePackageInvocation), printerHelp);
            _intentStrings.Add("PackagePrinter".ToLower(), "printer");
        }
        /// <summary>
        /// The constructor, taking the same arguments as any other.
        /// </summary>
        /// <param name="Args">Commandline arguments.</param>
        /// <param name="Intention">Context of the invocation.</param>
        /// <exception cref="Invocation.MissingArgumentException"></exception>
        public IntunePackageInvocation(string[] Args)
        {
            _intent = _intentStrings[Args[0].ToLower()];

            // Checks if flag is set.
            _createPackage = ParameterParser.GetFlag(Args, "-IntuneWinUtil");

            // Path can be null because IntuneWinUtil class will try to sort this later.
            if (ParameterParser.GetParameterValue(Args, "-IntuneWinUtil", out _intuneWinUtilPath))
            {
                if (!File.Exists(_intuneWinUtilPath))
                    throw new Invocation.InvalidNameOrPathException($"The provided IntuneWinAppUtil.exe path of {_intuneWinUtilPath} does not exist.");
            }

            if (_intent == "driver")
            {
                if (!ParameterParser.GetParameterValue(Args, "-Driver", out _driverName) && !ParameterParser.GetParameterValue(Args, "-Path", out _driverName))
                    throw new Invocation.MissingArgumentException("Invalid invocation: A valid driver name or path is required.");

                // Check if driver exists in system by attempting instantiation.
                if (_driverName.StartsWith(Invocation.SystemPath) || !File.Exists(_driverName))
                    _exportDriverFromSystem = true;
            }
            else
            {
                if (!ParameterParser.GetParameterValue(Args, "-PrinterName", out _printerName))
                    throw new Invocation.MissingArgumentException(
                        "Invalid invocation: A value must be provided for the PrinterName parameter."
                    );
            }

            // If no output path is provided, use a yet-to-be-created subdirectory of the current directory.
            if (!ParameterParser.GetParameterValue(Args, "-Output", out _outputPath))
                _outputPath = Path.Combine(Environment.CurrentDirectory, _packageName);
        }

        /// <summary>
        /// Switches based on the context and invokes the appropriate methods.
        /// </summary>
        /// <returns>The integer that Program.cs->Main() exits with.</returns>
        public int Invoke()
        {
            if (_intent == "printer")
                return PackagePrinter();
            else
                return PackageDriver();
        }
        /// <summary>
        /// Obtains a JSON definition of the specified printer and writes it to a file in the package folder.
        /// </summary>
        /// <returns>The full path of the configuration file.</returns>
        /// <exception cref="Invocation.MissingArgumentException"></exception>
        private string ExportPrinterDefinition()
        {
            string configPath;
            if (!_outputPath.EndsWith(".json"))
                configPath = Path.Combine(_outputPath, "config.json");
            else
                configPath = _outputPath;

            if (_printerName == null)
                throw new Invocation.MissingArgumentException("Invalid invocation: The PrinterName parameter is required when packaging a printer configuration.");

            var printerDefinition = Printer.SerializeExisting(_printerName);

            var parentDirectory = Path.GetDirectoryName(configPath);
            if (parentDirectory == null)
                throw new DirectoryNotFoundException($"Unable to find parent directory \"{parentDirectory}\" of \"{configPath}\".");

            FsHelper.CreateDirectory(parentDirectory);

            File.WriteAllText(configPath, printerDefinition);
            return configPath;
        }
        /// <summary>
        /// Invocation agnostic, copies the Printune.exe file into the package directory.
        /// </summary>
        /// <returns>Full path of the Printune.exe file copy.</returns>
        /// <exception cref="Invocation.InvalidNameOrPathException">Thrown when the destination is invalid.</exception>
        private string CopyPrintune()
        {
            string destination;
            if (_outputPath.EndsWith(".json"))
                destination = Path.GetDirectoryName(_outputPath);
            else
                destination = _outputPath;


            if (destination == null)
                throw new Invocation.InvalidNameOrPathException($"Invalid invocation: Output parameter requires a valid path.");

            var exeDirectory = Path.GetDirectoryName(Invocation.PrintuneExePath);

            var printuneFileName = new FileInfo(Invocation.PrintuneExePath).Name;

            var printuneDestination = Path.Combine(_outputPath, printuneFileName);

            var files = new List<string> {
                Invocation.PrintuneExePath,
                Path.Combine(exeDirectory, "Newtonsoft.Json.dll"),
                Path.Combine(exeDirectory, "System.CodeDom.dll"),
                Path.Combine(exeDirectory, "LICENSE"),
                Path.Combine(exeDirectory, "NewtonSoft.Json.LICENSE.md")
            };

            foreach (var file in files)
            {
                var path = Path.Combine(destination, Path.GetFileName(file));
                File.Copy(file, path, true);
            }
            
            Log.Write($"Printune executable copied from \"{Invocation.PrintuneExePath}\" to \"{printuneDestination}\".");
            return printuneDestination;
        }
        /// <summary>
        /// Invocation agnostic, uses IntuneWinAppUtil.exe to package the folder.
        /// </summary>
        /// <param name="Source">The source folder for the package.</param>
        /// <returns>The integer that Program.cs->Main() exits with.</returns>
        private int PackageFolder(string Source)
        {
            IntuneWinUtil.Result result;

            var printunePath = CopyPrintune();

            // We take the parent directory of the source folder
            // because that's where we're putting the .intunewin package.
            var destinationDir = Directory.GetParent(Source);
            if (destinationDir == null)
                throw new DirectoryNotFoundException($"Directory {Source} does not exist.");

            result = IntuneWinUtil.InvokeIntuneWinUtil(
                                            printunePath,
                                            Source, // Source
                                            destinationDir.FullName,
                                            _packageName,
                                            _intuneWinUtilPath
                                        );

            if (result.Success)
                Log.Write($"IntuneWinAppUtil.exe completed successfully with exit code {result.ExitCode}.");
            else
                Log.Write($"IntuneWinAppUtil.exe failed with exit code {result.ExitCode}.");

            Log.Write("OUTPUT");
            if (result.Success)
                Log.Write(result.Output, Indent: 1);
            else
                Log.Write(result.Output, true, 1);

            if (!string.IsNullOrEmpty(result.Error))
            {
                Log.Write("ERROR:");
                Log.Write(result.Error, true, 1);
            }

            if (result.Success)
                Log.Write($"Package created at {result.PackagePath}.");

            return result.ExitCode;
        }
        /// <summary>
        /// The printer context parent invocation of PackageFolder().
        /// </summary>
        /// <returns>The integer that Program.cs->Main() exits with.</returns>
        private int PackagePrinter()
        {
            var destination = CreatePackageFolder();
            var configFilePath = ExportPrinterDefinition();
            Log.Write($"Printer definition for \"{_printerName}\" created as \"{configFilePath}\".");

            if (_createPackage)
                return PackageFolder(destination);
            else
                return 0;
        }
        /// <summary>
        /// The driver context parent invocation of PackageFolder().
        /// </summary>
        /// <returns>The integer that Program.cs->Main() exits with.</returns>
        private int PackageDriver()
        {
            var destination = CreatePackageFolder();

            if (_driverName == null)
                throw new Invocation.MissingArgumentException("Invalid invocation: A valid DriverName value is required.");

            if (_exportDriverFromSystem)
            {
                PackageSystemDriver();
            }
            else
            {
                PackageDriverFromFile();
            }

            if (_createPackage)
                return PackageFolder(destination);
            else
                return 0;
        }
        /// <summary>
        /// Exports system driver for packaging.
        /// </summary>
        /// <exception cref="Invocation.MissingArgumentException"></exception>
        /// <exception cref="Invocation.InvalidNameOrPathException"></exception>
        /// <exception cref="Exception"></exception>
        private void PackageSystemDriver()
        {
            if (_driverName == null)
                throw new Invocation.MissingArgumentException("Invalid invocation: A valid DriverName value is required.");

            var driver = new PrinterDriver(_driverName);
            if (driver == null)
                throw new Invocation.InvalidNameOrPathException($"The driver \"{_driverName}\" does not exist on this system.");

            var result = driver.ExportDriver(_outputPath);

            if (result.Success)
                Log.Write($"Export of driver succeeded with exit code {result.ExitCode}:");
            else
                Log.Write($"Export of driver failed with exit code {result.ExitCode}:", true);

            Log.Write("OUTPUT:");
            Log.Write(result.Output, !result.Success, 1);

            if (!string.IsNullOrEmpty(result.Error))
            {
                Log.Write("ERROR:");
                Log.Write(result.Error, !result.Success, 1);
            }

            if (!result.Success)
                throw new Exception("PnpUtil.exe exited with an error.");
        }
        /// <summary>
        /// Copies a driver for packaging.
        /// </summary>
        /// <exception cref="Invocation.InvalidNameOrPathException"></exception>
        private void PackageDriverFromFile()
        {
            if (!File.Exists(_driverName) && !Directory.Exists(_driverName))
                throw new Invocation.InvalidNameOrPathException($"Invalid invocation: Driver path \"{_driverName}\" is not valid.");

            string driverDirectory = string.Empty;
            if (File.Exists(_driverName))
            {
                var dir = new FileInfo(_driverName).Directory;
                if (dir != null)
                    driverDirectory = dir.FullName;
                else
                    throw new Invocation.InvalidNameOrPathException($"Invalid invocation: Driver directory \"{driverDirectory}\" is not valid.");
            }

            var driverSubDirectories = Directory.GetDirectories(driverDirectory, "*", SearchOption.AllDirectories);
            var driverFiles = Directory.GetFiles(driverDirectory, "*", SearchOption.AllDirectories);

            foreach (string driverSubDirectory in driverSubDirectories)
            {
                var newPath = driverSubDirectory.Replace(driverDirectory, _outputPath);
                Directory.CreateDirectory(newPath);
            }

            foreach (string driverFile in driverFiles)
            {
                var newPath = driverFile.Replace(driverDirectory, _outputPath);
                File.Copy(driverFile, newPath, true);
            }
        }
        /// <summary>
        /// A wrapper for the FsHelper.CreateDirectory() method that recursively creates drectories.
        /// </summary>
        /// <returns>The full path of the target directory or null if failed.</returns>
        private string CreatePackageFolder()
        {
            var path = FsHelper.CreateDirectory(_outputPath);
            Log.Write($"Package directory \"{path}\" created.");
            return path;
        }
    }
}