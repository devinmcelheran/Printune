using System;
using System.Diagnostics;
using System.IO;

namespace Printune.Models
{
    public class PrinterPreference
    {
        public string PreferenceFile { get; private set; }

        public PrinterPreference(string preferenceFile)
        {
            PreferenceFile = ConfigReader.CopyToLocal(preferenceFile);
            if (!File.Exists(PreferenceFile))
                if (preferenceFile.StartsWith(Path.GetTempPath(), StringComparison.InvariantCultureIgnoreCase))
                    throw new FileNotFoundException("The specified printer preference file failed to copy to a local temporary file but did not throw an error.", preferenceFile);
                else
                    throw new FileNotFoundException("The specified printer preference file does not exist.", PreferenceFile);
        }

        public void Apply(string PrinterName)
        {
            if (!Invocation.IsElevated)
                throw new UnauthorizedAccessException("Applying printer preferences requires elevated (administrator) permissions.");

            using (var printer = Printer.FromExisting(PrinterName))
            {
                if (printer == null)
                    throw new ArgumentException($"The specified printer '{PrinterName}' does not exist.", nameof(PrinterName));

                PrintUIApply(PrinterName);
            }
        }

        private void PrintUIApply(string PrinterName)
        {
            var psi = new ProcessStartInfo()
            {
                FileName = "rundll32.exe",
                Arguments = $"printui.dll,PrintUIEntry /Ss /n \"{PrinterName}\" /a \"{PreferenceFile}\" /q",
                WindowStyle = ProcessWindowStyle.Hidden,
                WorkingDirectory = Directory.GetCurrentDirectory()
            };

            using (var proc = Process.Start(psi))
            {
                if (proc == null)
                    throw new InvalidOperationException("Failed to start the PrintUI process to apply printer preferences.");

                proc.WaitForExit();
                if (proc.ExitCode != 0)
                    throw new InvalidOperationException($"The PrintUI process to apply printer preferences exited with an error. Exit code: {proc.ExitCode}");

                Log.Write("Printing prefrences applied successfully.");
            }
        }
        public static void Export(string PrinterName, string OutputFile)
        {
            bool debug = false;
            #if DEBUG
            debug = true;
            #endif
            if (!Invocation.IsElevated && !debug)
                throw new UnauthorizedAccessException("Applying printer preferences requires elevated (administrator) permissions.");

            using (var printer = Printer.FromExisting(PrinterName))
            {
                if (printer == null)
                    throw new ArgumentException($"The specified printer '{PrinterName}' does not exist.", nameof(PrinterName));

                PrintUIExport(PrinterName, OutputFile);
            }
        }

        private static void PrintUIExport(string PrinterName, string OutputFile)
        {
            var psi = new ProcessStartInfo()
            {
                FileName = "rundll32.exe",
                Arguments = $"printui.dll,PrintUIEntry /Sr /n \"{PrinterName}\" /a \"{OutputFile}\" /q",
                WindowStyle = ProcessWindowStyle.Hidden,
                WorkingDirectory = Directory.GetCurrentDirectory()
            };

            using (var proc = Process.Start(psi))
            {
                if (proc == null)
                    throw new InvalidOperationException("Failed to start the PrintUI process to export printer preferences.");

                proc.WaitForExit();
                if (proc.ExitCode != 0)
                    Log.Write($"The PrintUI process to export printer preferences exited with an error. Exit code: {proc.ExitCode}", true);

                Log.Write($"Printing prefrences exported successfully to {OutputFile}");
            }
        }
    }
}