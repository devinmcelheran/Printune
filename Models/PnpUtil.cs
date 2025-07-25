using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Printune
{
    /// <summary>
    /// This class is a wrapper around the pnputil.exe tool.
    /// </summary>
    public static partial class PnpUtil
    {
        /// <summary>
        /// The pnputil is used for installing, uninstalling, and exporting drivers.
        /// Exporting drivers requires converting the driver name or files path to
        /// the Published Name, which is in the format "oem###.inf", drivers are
        /// enumerated to facilitate the mapping of names and paths to Published Names.
        /// </summary>
        private enum Intention
        {
            Install,
            Uninstall,
            Enumerate,
            Export
        };
        /// <summary>
        /// This is the path to the pnputil application.
        /// </summary>
        private static string _pnpUtilPath
        {
            get
            {
                var systemPath = Environment.GetFolderPath(Environment.SpecialFolder.System);
                return Path.Combine(systemPath, "pnputil.exe");
            }
        }
        /// <summary>
        /// Installs the provided driver *.inf file.
        /// </summary>
        /// <param name="FilePath">The path to the *.inf file.</param>
        /// <returns>PnpUtil.Result containing pnputil.exe stdout and stderr output, exit code, and commandline information.</returns>
        public static Result InstallDriver(string FilePath)
        {
            return InvokePnpUtil(Intention.Install, FilePath);
        }
        /// <summary>
        /// Uninstalls the provided driver *.inf file.
        /// </summary>
        /// <param name="FilePath">The path to the *.inf file.</param>
        /// <returns>PnpUtil.Result containing pnputil.exe stdout and stderr output, exit code, and commandline information.</returns>
        public static Result UninstallDriver(string FilePath)
        {
            return InvokePnpUtil(Intention.Uninstall, FilePath);
        }
        /// <summary>
        /// Invokes the pnputil.exe tool with the provided parameters.
        /// </summary>
        /// <param name="Intent">The intended operation (Install/Uninstall/Enumerate/Export)</param>
        /// <param name="Target">The driver being operated on, either file path, name, or null (if enumerating).</param>
        /// <param name="Destination">Export destination (only referenced when exporting).</param>
        /// <returns></returns>
        private static Result InvokePnpUtil(Intention Intent, string Target = null, string Destination = null)
        {
            Process proc = new Process();
            proc.StartInfo.FileName = _pnpUtilPath;

            switch (Intent)
            {
                case Intention.Install:
                    Log.Write($"Attempting installation of {Target}...");
                    proc.StartInfo.Arguments = $"/add-driver \"{Target}\" /install";
                    break;
                case Intention.Uninstall:
                    Log.Write($"Attempting uninstallation of {Target}...");
                    proc.StartInfo.Arguments = $"/delete-driver \"{Target}\"";
                    break;
                case Intention.Enumerate:
                    Log.Write($"Enumerating drivers...");
                    proc.StartInfo.Arguments = $"/enum-drivers";
                    break;
                case Intention.Export:
                    Log.Write($"Attempting export of {Target} to {Destination}.");
                    proc.StartInfo.Arguments = $"/export-driver \"{Target}\" \"{Destination}\"";
                    break;
            }

            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.RedirectStandardOutput = true;
            proc.StartInfo.RedirectStandardError = true;

            StringBuilder output = new StringBuilder();
            StringBuilder error = new StringBuilder();

            proc.OutputDataReceived += (sender, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    output.AppendLine(e.Data);
                }
            };

            proc.ErrorDataReceived += (sender, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    error.Append(e.Data);
                }
            };

            proc.Start();
            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();

            proc.WaitForExit();

            return new Result(proc.StartInfo.Arguments, proc.ExitCode, output.ToString(), error.ToString());
        }
        /// <summary>
        /// Exports Target driver to DestinationPath with option to specify DriverVersion.
        /// </summary>
        /// <param name="Target">The driver being exported.</param>
        /// <param name="DestinationPath">Folder in which driver should be exported.</param>
        /// <param name="DriverVersion">Optional driver version as string, can match pnputil date or version in DriverVersion property.</param>
        /// <returns>PnpUtil.Result containing stdout and stderr output, exit code, and command line information.</returns>
        /// <exception cref="Invocation.InvalidNameOrPathException"></exception>
        public static Result Export(string Target, string DestinationPath, string DriverVersion = null)
        {
            // Attempt to instantiate a PnpDriver object. Returns null if no such driver exists.
            PnpDriver oemTarget = PnpDriver.FromTarget(Target);

            // If null
            if (oemTarget == null)
                throw new Invocation.InvalidNameOrPathException($"[PnpDriver] could not be found with a target of \"{Target}\".");

            return InvokePnpUtil(Intention.Export, oemTarget.PublishedName, DestinationPath);
        }
        /// <summary>
        /// Invokes pnputil.exe to enumerate drivers and obtain output.
        /// </summary>
        /// <returns>PnpUtil.Result containining stdout and stderr output listing all drivers, exit code, and commandline information.</returns>
        private static Result EnumerateDrivers()
        {
            return InvokePnpUtil(Intention.Enumerate);
        }
        /// <summary>
        /// Obtains pnputil driver enumeration output and returns parsed result.
        /// </summary>
        /// <returns>IEnumerable<PnpDriver> containing enumerated drivers.</returns>
        /// <exception cref="Exception">Raised when pnputil.exe exits a non-zero return code.</exception>
        public static IEnumerable<PnpDriver> GetDrivers()
        {
            var result = EnumerateDrivers();

            if (!result.Success)
                throw new Exception($"An error occured, \"{result.CommandLine}\" exited with error code {result.ExitCode}.");

            return PnpDriver.FromPnpUtilOutput(result.Output);
        }
        /// <summary>
        /// A simple class containing stdout and stderr output, exit code, and commandline information captured from a pnputil.exe invocation.
        /// </summary>
        public class Result
        {
            /// <summary>
            /// The commandline parameters provided to pnputil.
            /// </summary>
            public readonly string CommandLine;
            /// <summary>
            /// The code that pnputil returned upon exit.
            /// </summary>
            public readonly int ExitCode;
            /// <summary>
            /// The stdout output gather when pnputil was run.
            /// </summary>
            public readonly string Output;
            /// <summary>
            /// The stderr output gathered when pnputil was run.
            /// </summary>
            public readonly string Error;
            /// <summary>
            /// Returns whether pnputil exited with a documented success return code.
            /// </summary>
            public bool Success { get
                {
                    return ExitCode == 0 // Success
                    || ExitCode == 259   // Already installed or newer version installed.
                    || ExitCode == 3010; // Reboot requirted.
                } }

            /// <summary>
            /// Constructs a result instance.
            /// </summary>
            /// <param name="CommandLineArguments">The arguments provided to pnputil when run (Process.StartInfo.Arguments).</param>
            /// <param name="ExitCode">The return code that pnputil exited with.</param>
            /// <param name="Output">The stdout output captured when pnputil was run.</param>
            /// <param name="Error">The stderr output capture when pnputil was run.</param>
            public Result(string CommandLineArguments, int ExitCode, string Output, string Error)
            {
                this.CommandLine = "pnputil.exe " + CommandLineArguments;
                this.ExitCode = ExitCode;
                this.Output = Output;
                this.Error = Error;
            }
        }
        /// <summary>
        /// A wrapper class that parses pnputil output into usable objects.
        /// </summary>
        public class PnpDriver
        {
            /// <summary>
            /// The installed driver name, taking the form "oem###.inf".
            /// </summary>
            public string PublishedName { get; private set; }
            /// <summary>
            /// The original *.inf file name for the driver.
            /// </summary>
            public string OriginalName { get; private set; }
            /// <summary>
            /// The manufacturer or driver provider.
            /// </summary>
            public string ProviderName { get; private set; }
            /// <summary>
            /// The device class name for which the driver applies.
            /// </summary>
            public string ClassName { get; private set; }
            /// <summary>
            /// The device class GUID for which the driver applies.
            /// </summary>
            public string ClassGuid { get; private set; }
            /// <summary>
            /// The version value extracted from the pnputil DriverVersion field.
            /// </summary>
            public Version DriverVersion { get; private set; }
            /// <summary>
            /// The date value extracted from the pnputil DriverVersion field.
            /// </summary>
            public DateTime DriverVersionDate { get; private set; }
            /// <summary>
            /// The name of the certificate used to sign the driver.
            /// </summary>
            public string SignerName { get; private set; }

            /// <summary>
            /// A basic constructor for the class.
            /// </summary>
            /// <param name="PublishedName">The installed driver name, taking the form "oem###.inf".</param>
            /// <param name="OriginalName">The original *.inf file name for the driver.</param>
            /// <param name="ProviderName">The manufacturer or driver provider.</param>
            /// <param name="ClassName">The device class name for which the driver applies.</param>
            /// <param name="ClassGuid">The device class GUID for which the driver applies.</param>
            /// <param name="DriverVersion">The version value extracted from the pnputil DriverVersion field.</param>
            /// <param name="DriverVersionDate">The date value extracted from the pnputil DriverVersion field.</param>
            /// <param name="SignerName">The name of the certificate used to sign the driver.</param>
            private PnpDriver(string PublishedName, string OriginalName, string ProviderName, string ClassName,
                            string ClassGuid, Version DriverVersion, DateTime DriverVersionDate, string SignerName)
            {
                this.PublishedName = PublishedName;
                this.OriginalName = OriginalName;
                this.ProviderName = ProviderName;
                this.ClassName = ClassName;
                this.ClassGuid = ClassGuid;
                this.DriverVersion = DriverVersion;
                this.DriverVersionDate = DriverVersionDate;
                this.SignerName = SignerName;
            }
            /// <summary>
            /// Parses and constructs an instance from a block of "pnputil /enum-drivers" output.
            /// </summary>
            /// <param name="PnpUtilTextBlock">A driver text block from running "pnputil /enum-driver".</param>
            /// <returns>PnpUtil.Driver representing the driver information found in the text block.</returns>
            private static PnpDriver FromPnpUtilBlock(string PnpUtilTextBlock)
            {
                var publishedNameMatch = Regex.Match(PnpUtilTextBlock, @"(?<=Published Name:\s+)(oem\d+\.inf)", RegexOptions.IgnoreCase);
                var originalNameMatch = Regex.Match(PnpUtilTextBlock, @"(?<=Original Name:\s+)(\S.+\.inf)", RegexOptions.IgnoreCase);
                var providerNameMatch = Regex.Match(PnpUtilTextBlock, @"(?<=Provider Name:\s+)(\S.+\S)", RegexOptions.IgnoreCase);
                var classNameMatch = Regex.Match(PnpUtilTextBlock, @"(?<=Class Name:\s+)(\S.+\S)", RegexOptions.IgnoreCase);
                var classGuidMatch = Regex.Match(PnpUtilTextBlock, @"(?<=Class GUID:\s+)(\S+)", RegexOptions.IgnoreCase);
                var driverVersionMatch = Regex.Match(PnpUtilTextBlock, @"(?<=Driver Version:\s+)(\S+)\s+(\S+\S)", RegexOptions.IgnoreCase);
                var signerNameMatch = Regex.Match(PnpUtilTextBlock, @"(?<=Signer Name:\s+)(\S.+\S)", RegexOptions.IgnoreCase);

                var versionDate = DateTime.ParseExact(driverVersionMatch.Groups[1].Value, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                var version = new Version(driverVersionMatch.Groups[2].Value);

                var driver = new PnpDriver
                (
                    publishedNameMatch.Groups[1].Value, // PublishedName
                    originalNameMatch.Groups[1].Value,  // OriginalName
                    providerNameMatch.Groups[1].Value,  // ProviderName
                    classNameMatch.Groups[1].Value,     // ClassName
                    classGuidMatch.Groups[1].Value,     // ClassGuid
                    version,                            // DriverVersion
                    versionDate,                        // DriverVersionDate
                    signerNameMatch.Groups[1].Value     // SignerName
                );

                return driver;
            }
            /// <summary>
            /// Takes "pnputil /enum-drivers" output, splits into individual driver blocks, runs them through the constructor, and returns them.
            /// </summary>
            /// <param name="PnpUtilOutput">Output from the "pnputil /enum-drivers" command.</param>
            /// <returns>List<PnpDriver> containing all enumerated drivers.</returns>
            public static List<PnpDriver> FromPnpUtilOutput(string PnpUtilOutput)
            {
                var drivers = new List<PnpDriver>();

                // I was having difficulty with Regex not splitting on the "\r\n" line endings,
                // so a pipe was inserted at each new driver block, which was then used for splitting.
                var blocks = PnpUtilOutput
                                .Replace("Microsoft PnP Utility", "")
                                .Replace("Published Name:", "|Published Name:")
                                .Split('|');

                foreach (var block in blocks)
                {
                    // Skip any null, empty, or blank lines.
                    if (string.IsNullOrEmpty(block.Trim()) || Regex.IsMatch(block.Trim(), "^$|\n\r"))
                        continue;

                    drivers.Add(FromPnpUtilBlock(block));
                }

                return drivers;
            }
            /// <summary>
            /// Returns the PnpDriver object for an installed driver.
            /// </summary>
            /// <param name="PublishedName">The "oem###.inf" name of the driver.</param>
            /// <param name="DriverVersion">Optional driver version.</param>
            /// <returns>PnpUtil.PnpDriver representing the driver, if found, or null, if not found, and most recent version unless specified.</returns>
            private static PnpDriver FromPublishedName(string PublishedName, string DriverVersion = null)
            {
                var drivers = PnpUtil.GetDrivers();

                if (DriverVersion == null)
                    return drivers
                            .Where(driver => driver.PublishedName == PublishedName)
                            .OrderByDescending(driver => driver.DriverVersion)
                            .FirstOrDefault();

                return drivers
                        .Where(driver => driver.PublishedName == PublishedName)
                        .Where(driver =>
                        {
                            if (driver.DriverVersion.ToString() == DriverVersion)
                                return true;


                            DateTime date;
                            try
                            {
                                date = DateTime.Parse(DriverVersion);
                                if (driver.DriverVersionDate == date)
                                    return true;
                            }
                            catch
                            {
                                return false;
                            }

                            return false;
                        })
                        .FirstOrDefault();
            }
            /// <summary>
            /// Returns the PnpDriver object for an installed driver.
            /// </summary>
            /// <param name="OriginalName">The original *.inf file name of the driver.</param>
            /// <param name="DriverVersion">Optional driver version.</param>
            /// <returns>PnpUtil.PnpDriver representing the driver, if found, or null, if not found, and most recent version unless specified.</returns>
            private static PnpDriver FromOriginalName(string OriginalName, string DriverVersion = null)
            {
                var fileName = Path.GetFileName(OriginalName);

                var drivers = PnpUtil.GetDrivers();

                if (DriverVersion == null)
                    return drivers
                            .Where(driver => driver.OriginalName == Path.GetFileName(OriginalName))
                            .OrderByDescending(driver => driver.DriverVersion)
                            .FirstOrDefault();

                return drivers
                        .Where(driver => driver.OriginalName == Path.GetFileName(OriginalName))
                        .Where(driver =>
                        {
                            if (driver.DriverVersion.ToString() == DriverVersion)
                                return true;


                            DateTime date;
                            try
                            {
                                date = DateTime.Parse(DriverVersion);
                                if (driver.DriverVersionDate == date)
                                    return true;
                            }
                            catch
                            {
                                return false;
                            }

                            return false;
                        })
                        .FirstOrDefault();
            }
            /// <summary>
            /// Returned PnpDriver from target.
            /// </summary>
            /// <param name="Target"></param>
            /// <param name="DriverVersion"></param>
            /// <returns></returns>
            public static PnpDriver FromTarget(string Target, string DriverVersion = null)
            {
                if (IsDriverStorePath(Target))
                {
                    var publishedName = DriverStorePathToPublishedName(Target);

                    if (publishedName != null)
                        return FromPublishedName(publishedName);
                    else
                        throw new Invocation.InvalidNameOrPathException(
                            "Invalid invocation: A valid driver published name or driver store path must be provided."
                            );
                }
                else if (Target.Trim().StartsWith("oem"))
                    return FromPublishedName(Target.Trim(), DriverVersion);
                else
                    return null;
            }
            private static bool IsDriverStorePath(string Target)
            {
                return Target.StartsWith(
                        Environment.GetFolderPath(Environment.SpecialFolder.System),
                        StringComparison.InvariantCultureIgnoreCase
                    );
            }
            private static string DriverStorePathToPublishedName(string Target)
            {
                if (!Invocation.IsElevated)
                    throw new Invocation.ElevatedPrivilegesRequiredException(
                                    "Error: Elevated privileges are required when exporting a driver using the Windows driver store path."
                                );

                var proc = new Process();

                proc.StartInfo.FileName = "dism.exe";
                proc.StartInfo.Arguments = $"/online /get-driverinfo /driver:{Target}";

                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.RedirectStandardError = true;

                StringBuilder output = new StringBuilder();

                proc.OutputDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        output.AppendLine(e.Data);
                    }
                };

                proc.Start();
                proc.BeginOutputReadLine();
                proc.WaitForExit();

                var publishedNameMatch = Regex.Match(output.ToString(), @"Published\sName\s+:\s+(oem\d+\.inf)");
                if (publishedNameMatch.Groups[1].Success)
                    return publishedNameMatch.Groups[1].Value;
                else
                    return null;
            }
        }
    }
}