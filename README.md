# Printune
A one-stop utility for creating Intune packages to deploy printers and printer drivers. The Printune project has both .Net 4.8 and 8 targets for convenience of both testing and deployment.

Printune can
- Package (using intunewinutil.exe)
- Add/Install
- Remove/Uninstall
- Verify (Detect)

Both network printers and drivers.

## Syntax
    printune.exe InstallDriver [-Recurse] {-Path <driver.inf> | -Path <folder\> } { -Name <driver name> } [-LogPath <file.log>]
    printune.exe UninstallDriver { -Driver <PrinterDriverName> | -Path <driver.inf> } [-LogPath <file.log>]
    
    printune.exe InstallPrinter { -PrinterName <PrinterName> } [ -Config <config.json> ] [-LogPath <file.log>]
    printune.exe UninstallPrinter { -PrinterName <PrinterName> } [-LogPath <file.log>]
    
    printune.exe PackageDriver { -Driver <PrinterDriverName> | -Path <driver.inf> } [ -Output <destination\> ] [ -IntuneWinUtil <path\intunewinutil.exe> ] [-LogPath <path\to\file.log>]
    printune.exe PackagePrinter { -PrinterName <PrinterName>} [ -Output <destination\> ] [ -IntuneWinUtil <path\intunewinutil.exe> ] [-LogPath <file.log>]
    
    printune.exe VerifyDriver -Name <DriverName> [-Version <Version>] [-LogPath <file.log>]
    printune.exe VerifyPrinter -Name <PrinterName> [-LogPath <file.log>]

    printune.exe Help
    printune.exe Help [-Template]

# Examples
## Logging
Logging is as simple as adding the `-LogPath` argument. Anything printed to the console will also be logged to the specified file.

    C:\> printune.exe <arguments> -LogPath C:\ProgramData\MyCompany\Printune\ThisPrinter.log

## Packaging
Printune can package a printer or printer driver, ready to deploy, using IntuneWinUtil.exe. But this isn't required. By optionally including the `-IntuneWinUtil` argument, the value being the path to the executable, you can have Printune create the package for you. If you would prefer to add some other contents (like a script) before packaging it, just omit the argument and find the package contents in the specified output folder.

    C:\> printune <arguments> -IntuneWinUtil $Home\Tools\IntuneWinUtil.exe
### Packaging a Printer
This is done by first configuring the printer on your own machine and then running this command. This will generate the necessary JSON configuration payload for the printer.

    C:\> printune.exe PackagePrinter -PrinterName "Finance Printer" -Output C:\Temp\FinancePrinter
### Packing a Driver
This can be done using the name of an installed driver, or the path of an .inf file somewhere in the filesystem, including the driverstore.

    C:> printune.exe PackageDriver -DriverName "Zerocks Global PCL6" -Output C:\Temp\FinancePrinter\
Or
    
    C:\> printune.exe PackageDriver -Path zerocks_uni_pcl6\x3UNIVZ.inf -Output C:\Temp\FinancePrinter\
## Detection
Printune can even be used as the detection method for printers and drivers. For drivers, there's also the added benefit checking the driver version.
### Detecting a Printer
    C:\> printune.exe VerifyPrinter -Name "Finance Printer"
### Detecting a Driver
    C:\> printune.exe VerifyDriver -Name "Zerocks Global PCL6" -Version "51055.300.0.0"
## Printers
Printune can install network printers using JSON files defining the printer.

The JSON template for defining printers (and printer ports) can be generated using the following command. Printune will accept printers as either a single object or an array in the JSON. This is helpful because you can maintain a single file listing all your printers. With this approach, you can reuse the same printer package and only change the install arguments for each deployment. There's no need to actually create many packages.

    C:\> printune Help -Template
    [
    {
        "PrinterName": "Hallway Printer",
        "Version": "1.0",
        "DriverName": "Zerocks Universal PCL",
        "DataType": "RAW",
        "PrintProcessor": "winprint",
        "Location": "In the hallway.",
        "PrinterPort": {
        "PrinterHostAddress": "1.2.3.4",
        "PortNumber": 9100,
        "SNMP": 1,
        "SNMPCommunity": "public"
        }
    },
    {
        "PrinterName": "Other Hallway Printer",
        "Version": "1.0",
        "DriverName": "Zerocks Universal PCL",
        "DataType": "RAW",
        "PrintProcessor": "winprint",
        "Location": "In the hallway.",
        "PrinterPort": {
        "PrinterHostAddress": "1.2.3.5",
        "PortNumber": 9100,
        "SNMP": 1,
        "SNMPCommunity": "public"
        }
    }
]
### Add a Printer Using Local File
You can specify a configuration file with the `-Config` argument, or Printune will check the current working directory for `config.json` by default.

    C:\> printune.exe InstallPrinter -Name FinancePrinter -Config FinancePrinter.json
### Add a Printer Using a Remote File
You can even specify a file that's accessible over HTTP/S.

    C:\> printune.exe InstallPrinter -Name FinancePrinter -Config https://print.corp.com/financeprinter.json
### Remove a Printer
    C:\> printune.exe UninstallPrinter -Name FinancePrinter
## Drivers
Printune can install and uninstall drivers with just a few arguments.
### Installing a Driver with a Single .inf File
    C:\> printune.exe InstallDriver -Path zerocks_uni_pcl6\x3UNIVZ.inf -Name "Zerocks Global PCL6"
### Installing a Driver with Many .inf Files
    C:\> printune.exe InstallDriver -Recurse -Path "driver\" -Name "Zerocks Global PCL6"
### Uninstalling a Driver Using an .inf File Path
    C:\> printune.exe UninstallDriver -Path "zerocks_uni_pcl6\x3UNIVZ.inf"
### Uninstalling a Driver by Name
    C:\> printune.exe UninstallDriver -Driver "Zerocks Global PCL6"

# Roadmap
This is a project that I've tinkered with to do away with the my collection of PowerShell scripts that we use for printer deployment. They work, but they have the occasional issue. This appears to be a more streamlined approach that even those less comfortable with PowerShell can use.

## Potential Future Features
- Automatically Generated Install/Uninstall/Detection Scripts
- Parameter Name Tweaks
- Better Help Text
    - Only Show Help Text Relevant to Attempted Operation
    - More Helpful Errors (And Remove Stack Trace)
        - Maybe Add a Command Switch to Enable Stack Trace Output

# Development
I'm a bit of an amateur developer. While I'm pretty hand in C#, my background is networking and systems adminsitration.

I'm open to any and all advice with respect to design, patterns, and project management.

## Other Clarifications
There's definitely some weirdness going on in this project. I started it as just a packaging tool that would use existing scripts. Because of that, I was sticking with .Net 8. But when I realized that managing printers and drivers on the endpoints with the same tool would be beneficial. I made those additions before realizing that it would mean deploying .Net 8 to all of our endpoints as well. This wasn't desirable, so I pivoted and added the `netstandard2.0` target (.Net 4.8).

This required the removal of several C# features (like static interface methods) as well as `System.Text.Json` in favor of `Newtonsoft.Json`. This introduced a lot of change, including many errors, some of which might not have been dealt with yet.