using System;
using System.IO;
using System.Linq;
using System.Management;

namespace Printune
{
    public class PrinterDriver : IDisposable
    {
        public string InstantiationString { get; private set; }
        private string _Name
        {
            get
            {
                if (_printerDriver != null)
                    return _printerDriver["Name"] as string;
                else
                    return null;
            }
        }
        private string _infPath
        {
            get
            {
                if (_printerDriver != null)
                    return _printerDriver["InfPath"] as string;
                else
                    return null;
            }
        }
        private string _provider
        {
            get
            {
                if (_printerDriver != null)
                    return _printerDriver["Provider"] as string;
                else
                    return null;
            }
        }
        public bool Exists
        {
            get
            {
                return _printerDriver != null;
            }
        }
        public string Version
        {
            get
            {

                if (_printerDriver["DriverVersion"] is UInt64 version)
                {
                    ushort rev = (ushort)(version & 0xFFFF);
                    ushort build = (ushort)((version >> 16) & 0xFFFF);
                    ushort minor = (ushort)((version >> 32) & 0xFFFF);
                    ushort major = (ushort)((version >> 48) & 0xFFFF);

                    return $"{major}.{minor}.{build}.{rev}";
                }
                return null;
            }
        }
        private ManagementObject _printerDriver;

        public PrinterDriver(string Driver, string Version = null)
        {
            InstantiationString = Driver;

            if (File.Exists(Path.GetFullPath(Driver)))
                _printerDriver = GetPrinterDriverCimByInf(Path.GetFullPath(Driver), Version);
            else
                _printerDriver = GetPrinterDriverCimByName(Driver, Version);
        }
        public void Remove()
        {
            if (_printerDriver != null)
            {
                _printerDriver.Delete();
                _printerDriver.Dispose();
            }
        }
        private static ManagementObject GetPrinterDriverCimByInf(string InfPath, string Version = null)
        {
            ManagementObjectSearcher searcher = null;

            try
            {
                searcher = new ManagementObjectSearcher(
                            "root/StandardCimv2",
                            $"SELECT * FROM MSFT_PrinterDriver"
                        );

                return searcher.Get()
                .Cast<ManagementObject>()
                .Where(pd =>
                {
                    if (String.IsNullOrEmpty(Version))
                        return pd["InfPath"] as string == InfPath;

                    using (var driver = new PrinterDriver(pd["InfPath"] as string))
                        return pd["InfPath"] as string == InfPath && driver.Version == Version;
                })
                .FirstOrDefault();
            }
            catch (Exception ex)
            {
                searcher?.Dispose();
                throw new Exception($"Failed to get the printer driver {InfPath} by inf path.", ex);
            }
        }
        private static ManagementObject GetPrinterDriverCimByName(string PrinterDriverName, string Version = null)
        {
            ManagementObjectSearcher searcher = null;
            try
            {
                searcher = new ManagementObjectSearcher("root/StandardCimv2",
                                $"SELECT * FROM MSFT_PrinterDriver WHERE Name = '{PrinterDriverName}'"
                                );

                var result = searcher.Get()
                            .Cast<ManagementObject>()
                            .Where(pd =>
                            {
                                if (String.IsNullOrEmpty(Version))
                                    return true;
                                
                                using (var driver = new PrinterDriver(pd["Name"] as string))
                                    return driver.Version == Version;
                            })
                            .FirstOrDefault();

                searcher.Dispose();
                return result;
            }
            catch (Exception ex)
            {
                searcher?.Dispose();
                throw new Exception($"Failed to get the printer driver {PrinterDriverName} by name.", ex);
            }
        }
        public static PrinterDriver EnablePrinterDriver(string PrinterDriverName)
        {
            ManagementClass printerDriverClass = null;
            ManagementBaseObject installParams = null;
            ManagementBaseObject result = null;

            try
            {
                printerDriverClass = new ManagementClass("ROOT/StandardCimv2:MSFT_PrinterDriver");
                installParams = printerDriverClass.GetMethodParameters("Add");
                installParams["Name"] = PrinterDriverName;

                result = printerDriverClass.InvokeMethod("Add", installParams, null);
                printerDriverClass.Dispose();
                installParams.Dispose();

                if (GetPrinterDriverCimByName(PrinterDriverName) == null)
                {
                    result.Dispose();
                    throw new Exception("Print driver enablement failed without throwing an error.");
                }
                else
                {
                    result.Dispose();
                    return new PrinterDriver(PrinterDriverName);
                }
            }
            catch (Exception ex)
            {
                printerDriverClass?.Dispose();
                installParams?.Dispose();
                result?.Dispose();
                throw new Exception($"Installation of the {PrinterDriverName} printer driver failed.", ex);
            }
        }
        public static bool DisablePrinterDriver(string PrinterDriverName)
        {
            var printerDriver = GetPrinterDriverCimByName(PrinterDriverName);

            // Return true if it doesn't exist,
            // that's as good as disabled.
            if (printerDriver == null)
            {
                return true;
            }
            else
            {
                printerDriver.Delete();
                printerDriver.Dispose();

                printerDriver = GetPrinterDriverCimByName(PrinterDriverName);
                if (printerDriver == null)
                    return true;
            }

            // If removing it doesn't work and doesn't throw an error,
            // just move on.
            printerDriver.Dispose();
            return false;
        }
        public PnpUtil.Result ExportDriver(string Destination)
        {
            if (!Exists || _infPath == null)
                throw new Invocation.InvalidNameOrPathException($"The [PrinterDriver] \"{InstantiationString}\" does not exist on the system.");

            return PnpUtil.Export(_infPath, Destination);
        }
        public void Dispose()
        {
            if (_printerDriver != null)
                _printerDriver.Dispose();
        }
    }
}