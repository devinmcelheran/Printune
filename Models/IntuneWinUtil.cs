using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Printune
{
    /// <summary>
    /// A wrapper object for the IntuneWinAppUtil.exe tool.
    /// </summary>
    public static class IntuneWinUtil
    {
        /// <summary>
        /// Searches current directory and %PATH% for "IntuneWinAppUtil.exe"
        /// and returns the copy with the highest file version.
        /// </summary>
        public static string IntuneWinUtilPath {
            get {
                return GetUtilPath();
            }
        }
        /// <summary>
        /// Searches current directory and %PATH% for "IntuneWinAppUtil.exe"
        /// and returns the copy with the highest file version.
        /// </summary>
        /// <returns>Full path to IntuneWinAppUtil.exe</returns>
        /// <exception cref="Invocation.InvalidNameOrPathException">Thrown when a copy of IntuneWinAppUtil.exe cannot be found.</exception>
        private static string GetUtilPath()
        {
            var envPath = Environment.GetEnvironmentVariable("PATH");
            List<string> searchLocations = new List<string>();

            if (envPath != null)
                searchLocations.AddRange(envPath.Split(';'));

            searchLocations.Add(Environment.CurrentDirectory);

            List<string> intuneWinUtilCopies = new List<string>();

            foreach (var location in searchLocations)
            {
                var results = Directory.GetFiles(location, "IntuneWinAppUtil.exe");

                if (results != null)
                    intuneWinUtilCopies.AddRange(results);
            }

            if (intuneWinUtilCopies.Count == 0)
            {
                throw new Invocation.InvalidNameOrPathException("Unable to find 'IntuneWinAppUtil.exe'."
                    + " Please provide a path using the 'IntuneWinUtil' parameter or place it in directory in %PATH%.");
            }

            // Sort according to version and return the greatest.
            return intuneWinUtilCopies
                        .Select(file => FileVersionInfo.GetVersionInfo(file))
                        .OrderByDescending(file => file.FileVersion)
                        .First()
                        .FileName;
        }
        /// <summary>
        /// The primary command wrapper for the IntuneWinAppUtil.exe tool.
        /// </summary>
        /// <param name="SetupFileName">The name of the setup file within the package.</param>
        /// <param name="Source">The containing folder that is to be packaged.</param>
        /// <param name="Destination">The destination for the package file.</param>
        /// <param name="PackageName">The name of the package.</param>
        /// <param name="UtilPath">Optional user-provided path.</param>
        /// <returns></returns>
        /// <exception cref="Invocation.InvalidNameOrPathException"></exception>
        public static Result InvokeIntuneWinUtil(string SetupFileName, string Source, string Destination, string PackageName, string UtilPath = null)
        {
            // Set UtilPath - should throw an error if none are found.
            if (UtilPath == null)
                UtilPath = IntuneWinUtilPath;

            // If no error is thrown and UtilPath is still null, throw an error.
            if (UtilPath == null)
                throw new Invocation.InvalidNameOrPathException("Unable to locate IntuneWinAppUtil.exe in current directory or %PATH%.");

            Process proc = new Process();

            proc.StartInfo.FileName = UtilPath;
            proc.StartInfo.Arguments = $"-c {Source} -o {Destination} -s {SetupFileName} -q";

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

            var packagePath = Path.Combine(Destination, "printune.intunewin");
            var newPath = packagePath.Replace("printune.intunewin", $"{PackageName}.intunewin");

            if (File.Exists(newPath))
                newPath = newPath.Replace(".intunewin", $"_{DateTime.Now.ToString("yyyyMMdd-HHmm")}.intunewin");
            if (File.Exists(packagePath))
                File.Move(packagePath, newPath);

            return new Result(proc.StartInfo.Arguments, proc.ExitCode, output.ToString(), error.ToString());
        }
        /// <summary>
        /// A simple class to encapsulate the IntuneWinAppUtil.exe results.
        /// </summary>
        public class Result
        {
            /// <summary>
            /// The command that was run.
            /// </summary>
            public readonly string CommandLine;
            /// <summary>
            /// The exit code of the command.
            /// </summary>
            public readonly int ExitCode;
            /// <summary>
            /// The stdout output of the command.
            /// </summary>
            public readonly string Output;
            /// <summary>
            /// The stderr output of the command.
            /// </summary>
            public readonly string Error;
            /// <summary>
            /// Whether or not the command was successful.
            /// </summary>
            public bool Success { get { return ExitCode == 0; } }

            /// <summary>
            /// Creates an IntuneWinAppUtil.Result object.
            /// </summary>
            /// <param name="CommandLine">The arguments fed to the command.</param>
            /// <param name="ExitCode">The code that the command exited with.</param>
            /// <param name="Output">The stdout output.</param>
            /// <param name="Error">The stderr output.</param>
            public Result(string CommandLine, int ExitCode, string Output, string Error)
            {
                this.CommandLine = "IntuneWinAppUtil.exe " + CommandLine;
                this.ExitCode = ExitCode;
                this.Output = Output;
                this.Error = Error;
            }
        }
    }
}