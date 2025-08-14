
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Runtime.Versioning;
using Microsoft.Win32;

namespace Printune.Models
{
    public class PrintProcessor
    {
        private static string PrintProcessorRegistryPath => $@"SYSTEM\CurrentControlSet\Control\Print\Environments\Windows {GetSystemArchitecture()}\Print Processors\";
        public string Name;
        public string RegistryKey => $@"HKEY_LOCAL_MACHINE\{PrintProcessorRegistryPath}\{Name}";

        public PrintProcessor()
        {
            throw new System.NotImplementedException("This constructor is not meant to be used directly.");
        }

        private PrintProcessor(string name)
        {
            Name = name;
        }

        public static List<PrintProcessor> GetAllPrintProcessors()
        {
            var printProcessors = Registry.LocalMachine.OpenSubKey(PrintProcessorRegistryPath).GetSubKeyNames();
            var result = new List<PrintProcessor>();
            foreach (var pp in printProcessors)
            {
                result.Add(new PrintProcessor(pp));
            }
            return result;
        }

        public static bool Exists(string name)
        {
            var exists = GetAllPrintProcessors().Select(pp => pp.Name).ToList().Contains(name, System.StringComparer.InvariantCultureIgnoreCase);
            if (!exists)
            {
                var registryKey = $@"HKEY_LOCAL_MACHINE\{PrintProcessorRegistryPath}\{name}";
                Log.Write($"Print processor {name} does not exist at {registryKey}", true);
            }
            return exists;
        }

        public static string GetSystemArchitecture()
        {
            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor"))
            {
                foreach (var item in searcher.Get())
                {
                    switch (item["Architecture"].ToString())
                    {
                        case "0": return "NT x86";
                        case "9": return "x64";
                        case "12": return "ARM64";
                        default: throw new System.NotSupportedException($"Unknown system architecture ({item["Architecture"]}) of CPU {item["Name"]} ({item["Description"]}).");
                    }
                }
                throw new System.NotSupportedException($"Unknown system architecture.");
            }
        }
    }
}