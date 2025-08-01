using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;

namespace Printune
{
    /// <summary>
    /// A wrapper class for the Win32_Printer CIM class.
    /// </summary>
    public class Printer : IDisposable
    {
        [JsonProperty("PrinterName")]
        public string Name { get; set; } = string.Empty;
        public string Location { get; set; } = string.Empty;
        public string Version { get; set; } = string.Empty;
        public string DriverName { get; set; } = string.Empty;
        public string DataType { get; set; } = string.Empty;
        public string PrintProcessor { get; set; } = string.Empty;
        [JsonIgnore]
        public string PortName
        {
            get
            {
                if (_printerPort != null)
                    return _printerPort.Name;
                else
                    return null;
            }
        }
        [JsonIgnore]
        public bool Exists
        {
            get
            {
                using (ManagementObject printer = GetPrinterManagementObject(Name))
                {
                    return printer != null;
                }
            }
        }
        [JsonProperty("PrinterPort")]
        public PrinterPort _printerPort { get; set; }

        public Printer() { }
        public static Printer FromConfig(string PrinterName, string ConfigContent)
        {
            var printerList = new List<Printer>();

            JsonLoadSettings loadSettings = new JsonLoadSettings
            {
                CommentHandling = CommentHandling.Ignore
            };

            JsonSerializerSettings serializerSettings = new JsonSerializerSettings
            {
                ContractResolver = new CaseInsensitiveContractResolver()
            };

            JObject config = JObject.Parse(ConfigContent, loadSettings);

            if (config.Root.Type == JTokenType.Array)
            {
                foreach (var item in config.Root.Values())
                {
                    try
                    {
                        var printer = JsonConvert.DeserializeObject<Printer>(item.ToString(), serializerSettings);
                        if (printer == null)
                            throw new Invocation.ConfigurationFileException("An error occurred while parsing the provided configuration file.");

                        printerList.Add(printer);
                    }
                    catch (Invocation.ConfigurationFileException)
                    {
                        throw;
                    }
                    catch (System.Exception ex)
                    {
                        throw new Invocation.ConfigurationFileException("An error occurred while parsing the provided configuration file.", ex);
                    }
                }
            }
            else
            {
                var printer = JsonConvert.DeserializeObject<Printer>(config.Root.ToString(), serializerSettings);
                if (printer == null)
                    throw new Invocation.ConfigurationFileException("An error occurred while parsing the provided configuration file.");

                printerList.Add(printer);
            }

            if (printerList.Where(printer => printer.DriverName == PrinterName).Count() > 1)
                throw new Invocation.ConfigurationFileException($"Invalid configuration file contains more than one printer definition name '{PrinterName}'.");

            try
            {
                return printerList.Where(printer => printer.Name == PrinterName).First();
            }
            catch (InvalidOperationException)
            {
                throw new Invocation.InvalidNameOrPathException($"The provided configuration file does not contain a printer named '{PrinterName}'.");
            }
        }
        public static Printer FromExisting(string PrinterName)
        {
            var printerManagementObject = GetPrinterManagementObject(PrinterName);

            if (printerManagementObject == null)
                return null;

            var printer = new Printer
            {
                Name = printerManagementObject[nameof(Name)] as string ?? string.Empty,
                Location = printerManagementObject[nameof(Location)] as string ?? string.Empty,
                DriverName = printerManagementObject[nameof(DriverName)] as string ?? string.Empty,
                DataType = printerManagementObject[nameof(DataType)] as string ?? string.Empty,
                PrintProcessor = printerManagementObject[nameof(PrintProcessor)] as string ?? string.Empty
            };

            var printerManagementObjectPortName = printerManagementObject[nameof(PortName)] as string;
            if (printerManagementObjectPortName == null)
                throw new Exception($"CIM instance of [Printer] named {PrinterName} returned null PortName value.");

            printer._printerPort = PrinterPort.FromExisting(printerManagementObjectPortName);

            printerManagementObject.Dispose();
            return printer;
        }
        public ManagementObject GetPrinterManagementObject()
        {
            if (PortName != null)
                return GetPrinterManagementObject(Name);
            else
                return null;
        }
        public static ManagementObject GetPrinterManagementObject(string PrinterName)
        {
            var scope = "root/StandardCimv2";
            var query = $"SELECT * FROM MSFT_Printer WHERE NAME = '{PrinterName}'";

            var printerManagementObject = new ManagementObjectSearcher(scope, query)
                    .Get()
                    .Cast<ManagementObject>()
                    .FirstOrDefault();

            return printerManagementObject;
        }
        public static List<ManagementObject> GetAllPrinterManagementObject()
        {
            var scope = "root/StandardCimv2";
            var query = $"SELECT * FROM MSFT_Printer";

            return new ManagementObjectSearcher(scope, query)
                    .Get()
                    .Cast<ManagementObject>()
                    .ToList();
        }
        public void Remove()
        {
            try
            {
                ManagementObject printer = GetPrinterManagementObject();
                if (printer != null)
                {
                    printer.Delete();

                    if (this.Exists)
                    {
                        printer.Dispose();
                        throw new Exception($"Removing [Printer] named {Name} failed without throwing an error.");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error occurred while removing [Printer] named \"{Name}\".", ex);
            }

            try
            {
                if (_printerPort != null)
                    _printerPort.Remove();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error occurred while removing [PrinterPort] named \"{PortName}\".", ex);
            }

            // We don't want to attempt disabling the driver
            // if it's still in use by other printers.
            var printers = GetAllPrinterManagementObject();
            bool driverInUse = printers
                                    .Where(p => DriverName == p[nameof(DriverName)] as string)
                                    .Count() > 0;

            printers.ForEach(p => p.Dispose());

            if (driverInUse)
            {
                Log.Write($"[PrinterDriver] named \"{DriverName}\" still in use by other printers.");
                return;
            }

            if (PrinterDriver.DisablePrinterDriver(DriverName))
                Log.Write($"[PrinterDriver] named \"{DriverName}\" was disabled successfully.");
            else
                Log.Write($"[PrinterDriver] named \"{DriverName}\" could not be disabled.", true);
        }
        private bool ChangesPending()
        {
            if (Exists == false)
                return true;

            bool _updatePending = false;

            var printerManagementObject = GetPrinterManagementObject();

            if (printerManagementObject == null)
                throw new Invocation.InvalidNameOrPathException($"The printer {Name} could not be found on the system.");

            _updatePending |= Name != printerManagementObject[nameof(Name)] as string;
            _updatePending |= Location != printerManagementObject[nameof(Location)] as string;
            _updatePending |= DriverName != printerManagementObject[nameof(DriverName)] as string;
            _updatePending |= DataType != printerManagementObject[nameof(DataType)] as string;
            _updatePending |= PrintProcessor != printerManagementObject[nameof(PrintProcessor)] as string;

            printerManagementObject.Dispose();
            return _updatePending;
        }
        public void Commit()
        {
            if (!ChangesPending()) return;

            if (Name == null)
                throw new ArgumentNullException("Unable to update [Printer] object with null 'PrinterName' property.");

            ManagementObject printer = GetPrinterManagementObject();
            if (printer != null)
            {
                printer.SetPropertyValue(nameof(Name), Name);
                printer.SetPropertyValue(nameof(Location), Location);
                printer.SetPropertyValue(nameof(DriverName), DriverName);
                printer.SetPropertyValue(nameof(DataType), DataType);
                printer.SetPropertyValue(nameof(PrintProcessor), PrintProcessor);

                if (_printerPort != null)
                    _printerPort.Commit();

                printer.Put();
                printer.Dispose();
            }
            else
            {
                if (_printerPort == null)
                    throw new Exception("[Printer] object somehow instantiated without a [PrinterPort], this shouldn't happen...");

                if (!_printerPort.Exists)
                    _printerPort.Commit();

                var cimMethod = "AddByExistingPort";
                var printerClass = new ManagementClass("root/StandardCimv2", "MSFT_Printer", null);
                var printerParameters = printerClass.GetMethodParameters(cimMethod);

                printerParameters[nameof(Name)] = Name;
                printerParameters[nameof(Location)] = Location;
                printerParameters[nameof(DriverName)] = DriverName;
                printerParameters[nameof(DataType)] = DataType;
                printerParameters[nameof(PrintProcessor)] = PrintProcessor;
                printerParameters[nameof(PortName)] = PortName;

                printerClass.InvokeMethod(cimMethod, printerParameters, null);
            }
        }
        public void OnDeserialized()
        {
            if (string.IsNullOrEmpty(Name))
                throw new ArgumentNullException("Unable to create printer port with null value for PrinterName property.");

            if (string.IsNullOrEmpty(DriverName))
                throw new ArgumentNullException("Unable to create printer port with null value for DriverName property.");

            if (string.IsNullOrEmpty(DataType))
                throw new ArgumentNullException("Unable to create printer port with null value for DataType property.");

            if (string.IsNullOrEmpty(PrintProcessor))
                throw new ArgumentNullException("Unable to create printer port with null value for PrintProcessor property.");

            // _printerManagementObject = GetPrinterManagementObject(Name!);

            // // If the value returned is null, then create a new one.
            // if (_printerManagementObject == null)
            //     _printerManagementObject = new ManagementClass("root/StandardCimv2", "MSFT_Printer", null).CreateInstance();
        }
        public string Serialize()
        {
            return JsonConvert.SerializeObject(this, Formatting.Indented);
        }
        public static string SerializeExisting(string PrinterName)
        {
            var printer = FromExisting(PrinterName);
            if (printer == null)
                return null;
            else
                return printer.Serialize();
        }

        public void Dispose()
        {
            _printerPort.Dispose();
        }
    }
}