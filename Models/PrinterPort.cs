using System;
using System.Linq;
using System.Management;
using Newtonsoft.Json;

namespace Printune
{

    public class PrinterPort : IDisposable
    {
        [JsonProperty("PortName")]
        public string Name { get; set;}
        public string HostAddress { get; set; }
        public UInt32 PortNumber { get; set; } = 0;
        public UInt32 Protocol { get; set; } = 1;
        public string Caption { get; set; }
        public bool ByteCount { get; set; } = false;
        public string Queue { get; set; }
        public string SNMPCommunity { get; set; }
        public bool SNMPEnabled { get; set; } = false;
        public UInt32 SNMPDevIndex { get; set; }
        [JsonIgnore]
        private ManagementObject _printerPortManagementObject;
        [JsonIgnore]
        public bool UpdatePending { get { return ChangesPending(); } }
        [JsonIgnore]
        public bool Exists
        {
            get
            {
                if (Name == null)
                    return false;
            
                ManagementObject printerPort = GetPrinterPortManagementObject(Name);
                return printerPort != null;
            }
        }

        public PrinterPort() {}
        private static ManagementObject GetPrinterPortManagementObject(string PrinterPortName)
        {
            var query = $"SELECT * FROM Win32_TcpIpPrinterPort WHERE NAME = '{PrinterPortName}'";

            var printerPortManagementObject = new ManagementObjectSearcher(query)
                    .Get()
                    .Cast<ManagementObject>()
                    .FirstOrDefault();

            return printerPortManagementObject;
        }
        public static PrinterPort FromExisting(string PrinterName)
        {
            var printerPortManagementObject = GetPrinterPortManagementObject(PrinterName);

            if (printerPortManagementObject == null)
                return null;
            else
                return FromExisting(printerPortManagementObject);
        }
        public static PrinterPort FromExisting(ManagementObject CimInstance)
        {
            var printerPort = new PrinterPort();
            printerPort._printerPortManagementObject = CimInstance;

            printerPort.Name = (string) printerPort._printerPortManagementObject[nameof(Name)];
            printerPort.Protocol = (UInt32) printerPort._printerPortManagementObject[nameof(Protocol)];
            printerPort.Caption = (string) printerPort._printerPortManagementObject[nameof(Caption)];
            printerPort.Queue = (string) printerPort._printerPortManagementObject[nameof(Queue)];
            printerPort.PortNumber = (UInt32) printerPort._printerPortManagementObject[nameof(PortNumber)];
            printerPort.HostAddress = (string) printerPort._printerPortManagementObject[nameof(HostAddress)];
            printerPort.SNMPCommunity = (string) printerPort._printerPortManagementObject[nameof(SNMPCommunity)];
            printerPort.SNMPEnabled = (bool) printerPort._printerPortManagementObject[nameof(SNMPEnabled)];
            if (printerPort.SNMPEnabled)
                printerPort.SNMPDevIndex = (UInt32) printerPort._printerPortManagementObject[nameof(SNMPDevIndex)];
            if (printerPort._printerPortManagementObject[nameof(ByteCount)] != null)
                printerPort.ByteCount = (bool) printerPort._printerPortManagementObject[nameof(ByteCount)];

            return printerPort;
        }
        public void Remove()
        {
            if (_printerPortManagementObject != null)
            {
                _printerPortManagementObject.Delete();       
                _printerPortManagementObject.Dispose();
            }
        }
        private bool ChangesPending()
        {
            if (_printerPortManagementObject == null)
                return false;

            bool _updatePending = false;

            _updatePending |= Name != _printerPortManagementObject[nameof(Name)] as string;
            _updatePending |= Caption != _printerPortManagementObject[nameof(Caption)] as string;
            _updatePending |= Queue != _printerPortManagementObject[nameof(Queue)] as string;
            _updatePending |= HostAddress != _printerPortManagementObject[nameof(HostAddress)] as string;
            _updatePending |= SNMPCommunity != _printerPortManagementObject[nameof(SNMPCommunity)] as string;

            // These properties (int and bool) are not nullable, they throw errors when compared
            // to null. The ManagementObject will return null if the value is not set, which breaks
            // the comparison, so null checking is required before comparison.

            if (_printerPortManagementObject[nameof(Protocol)] !=null)
                _updatePending |= Protocol != (UInt32) _printerPortManagementObject[nameof(Protocol)];

            if (_printerPortManagementObject[nameof(ByteCount)] != null)
                _updatePending |= ByteCount != (bool) _printerPortManagementObject[nameof(ByteCount)];

            if(_printerPortManagementObject[nameof(SNMPEnabled)] != null)
                _updatePending |= SNMPEnabled != (bool) _printerPortManagementObject[nameof(SNMPEnabled)];
        
            if (_printerPortManagementObject[nameof(PortNumber)] != null)
                _updatePending |= PortNumber != (UInt32) _printerPortManagementObject[nameof(PortNumber)];

            if (_printerPortManagementObject[nameof(SNMPDevIndex)] != null)
                _updatePending |= SNMPDevIndex != (uint) _printerPortManagementObject[nameof(SNMPDevIndex)];

            return _updatePending;
        }
        public void Commit()
        {
            if (Name == null)
                throw new ArgumentNullException("Unable to update [PrinterPort] object with null 'Name' property.");

            if (Exists)
                _printerPortManagementObject = GetPrinterPortManagementObject(Name);
            else
                using (var managementClass = new ManagementClass("Win32_TcpIpPrinterPort"))
                    _printerPortManagementObject = managementClass.CreateInstance();
        
            if (!ChangesPending()) return;

            if (Name != null)
                _printerPortManagementObject[nameof(Name)] = Name;

            _printerPortManagementObject[nameof(Protocol)] = Protocol;

            if (Caption != null)
                _printerPortManagementObject[nameof(Caption)] = Caption;
            if (ByteCount != false)
                _printerPortManagementObject[nameof(ByteCount)] = ByteCount;
            if (Queue != null)
                _printerPortManagementObject[nameof(Queue)] = Queue;
            if (PortNumber != 0)
                _printerPortManagementObject[nameof(PortNumber)] = PortNumber;
            if (HostAddress != null)
                _printerPortManagementObject[nameof(HostAddress)] = HostAddress;
            if (SNMPCommunity != null)
                _printerPortManagementObject[nameof(SNMPCommunity)] = SNMPCommunity;
            if (SNMPEnabled != false)
                _printerPortManagementObject[nameof(SNMPEnabled)] = SNMPEnabled;
            if (SNMPDevIndex != 0)
                _printerPortManagementObject[nameof(SNMPDevIndex)] = SNMPDevIndex;

            _printerPortManagementObject.Put();
        }
        public void OnDeserialized()
        {
            if (Name == null)
                throw new ArgumentNullException("Unable to create printer port with null value for Name property.");

            if (PortNumber == 0)
                throw new ArgumentNullException("Unable to create printer port with null value for PortNumber property.");

            if (HostAddress == null)
                throw new ArgumentNullException("Unable to create printer port with null value for PrinterHostAddress property.");

            _printerPortManagementObject = GetPrinterPortManagementObject(Name);

            // If the value returned is null, then create a new one.
            if (_printerPortManagementObject == null)
                _printerPortManagementObject = new ManagementClass("Win32_TcpIpPrinterPort").CreateInstance();
        }
        public void Dispose()
        {
            if (_printerPortManagementObject != null)
                _printerPortManagementObject.Dispose();
        }
    }
}