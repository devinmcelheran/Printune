using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Principal;

namespace Printune
{
    /// <summary>
    /// A static class for instantiating IInvocation objects.
    /// </summary>
    public static class Invocation
    {
        public static string[] Args
        {
            get
            {
                return Environment.GetCommandLineArgs().Skip(1).ToArray();
            }
        }
        /// <summary>
        /// The different contexts in which Printune.exe can be run.
        /// </summary>
        public enum Context
        {
            Help,
            InstallDriver,
            UninstallDriver,
            PackageDriver,
            InstallPrinter,
            UninstallPrinter,
            PackagePrinter
        }
        /// <summary>
        /// A dictionary that maps the context to the appropriate IInvocationContext implementation.
        /// </summary>
        private static Dictionary<string, Type> ContextMap = new Dictionary<string, Type>();
        private static Dictionary<string, string> ContextHelpMap = new Dictionary<string, string>();
        /// <summary>
        /// Registers the context mappings.
        /// </summary>
        public static void RegisterContext(string context, Type type, string helpMessage)
        {
            ContextMap.Add(context, type);
            ContextHelpMap.Add(context, helpMessage);
        }
        public static string HelpText
        {
            get
            {
                string helpText = string.Empty;
                var contexts = ContextMap.Keys.ToList();
                contexts.Sort();
                foreach (var context in contexts)
                {
                    helpText += $"{ContextHelpMap[context]}\n";
                }
                return helpText;
            }
        }
        /// <summary>
        /// A simple, less verbose, wrapper for referencing the system path.
        /// </summary>
        public static string SystemPath {
            get {
                return Environment.GetFolderPath(Environment.SpecialFolder.System);
            }
        }
        /// <summary>
        /// A simple, less verbose wrapper for determining if the invoking user has administrative privileges or not.
        /// </summary>
        public static bool IsElevated
        {
            get
            {
                return new WindowsPrincipal(WindowsIdentity.GetCurrent())
                            .IsInRole(WindowsBuiltInRole.Administrator);
            }
        }
        
        public static bool RunningDebug
        {
            get
            {
                #if DEBUG
                return true;
                #endif
                #pragma warning disable CS0162 // Unreachable code detected
                return false;
            }
        }
        /// <summary>
        /// The path of the current executable, used for when copying it into a package.
        /// </summary>
        public static string PrintuneExePath
        {
            get
            {
                var process = System.Diagnostics.Process.GetCurrentProcess();
                var module = process.MainModule;

                string path;
                if (module != null)
                    path = module.FileName;
                else
                    path = string.Empty;

                // If running in debug, use the "printunepath" environment variable that's configured in VS Code launch.json.
#if DEBUG
                path = Environment.GetEnvironmentVariable("printunepath");
#endif

                return path;
            }
        }
        /// <summary>
        /// The path to the directory containing the Printune executable.
        /// </summary>
        public static string PrintuneDir => Directory.GetParent(PrintuneExePath).FullName;
        /// <summary>
        /// The path to the paramters file.
        /// </summary>
        public static string ParamFile
        {
            get
            {
#if DEBUG
                return @"C:\temp\B38U\parameters.json";
#endif
                return Path.Combine(PrintuneDir, "parameters.json");
            }
        }
        /// <summary>
        /// Gets the Context enum from a string in a case-insensitive way and fails to help context.
        /// </summary>
        /// <param name="CommandContext">String containing the invocation context.</param>
        /// <returns></returns>
        private static Context GetContext(string CommandContext)
        {
            Context invocationContext;
            // true -> case-insensitive match
            if (Enum.TryParse<Context>(CommandContext, true, out invocationContext))
                return invocationContext;
            else
                return Context.Help;
        }
        /// <summary>
        /// Returns an instance of the appropriate invocation context.
        /// </summary>
        /// <param name="Args">Commandline arguments.</param>
        /// <returns>IInvocationContext to be invoked.</returns>
        public static IInvocationContext Parse(string[] Args)
        {
            if (Args.Count() == 0)
                return new HelpInvocation();

            Context context;

            try
            {
                context = GetContext(Args[0]);
            }
            catch
            {
                return new HelpInvocation($"Invalid invocation context '{Args[0]}'.");
            }

            Type contextType = ContextMap[Args[0].ToLower()];

            if (contextType == null)
                return new HelpInvocation($"Invalid parameter: {Args[0]}");

            var constructor = contextType.GetConstructor(new[] { typeof(string[]) });

            if (constructor == null)
                throw new InvalidOperationException($"No constructor found for context {contextType.Name} with paramter of type string[].");

            return (IInvocationContext)constructor.Invoke(new object[] { Args });
        }
        // Here onward are the various invocation errors
        // that inform the HelpInvocation as to what
        // help information should be displayed.
        public class MissingArgumentException : Exception
        {
            public MissingArgumentException() : base() { }
            public MissingArgumentException(string message) : base(message) { }
            public MissingArgumentException(string message, Exception innerException) : base(message, innerException) { }
        }
        public class InvalidNameOrPathException : Exception
        {
            public InvalidNameOrPathException() : base() { }
            public InvalidNameOrPathException(string message) : base(message) { }
            public InvalidNameOrPathException(string message, Exception innerException) : base(message, innerException) { }
        }
        public class ConfigurationFileException : Exception
        {
            public ConfigurationFileException() : base() { }
            public ConfigurationFileException(string message) : base(message) { }
            public ConfigurationFileException(string message, Exception innerException) : base(message, innerException) { }
        }
        public class ExternalDependencyNotFoundException : Exception
        {
            public ExternalDependencyNotFoundException() : base() { }
            public ExternalDependencyNotFoundException(string message) : base(message) { }
            public ExternalDependencyNotFoundException(string message, Exception innerException) : base(message, innerException) { }
        }

        internal class ElevatedPrivilegesRequiredException : Exception
        {
            public ElevatedPrivilegesRequiredException() { }
            public ElevatedPrivilegesRequiredException(string message) : base(message) { }
            public ElevatedPrivilegesRequiredException(string message, Exception innerException) : base(message, innerException) { }
        }
    }
}