using System;
using System.Collections.Generic;

namespace Printune
{
    public static class ParameterParser
    {
        public static bool GetParameterValue(string ParameterName, out string Value)
        {
            return GetParameterValue(Invocation.Args, ParameterName, out Value);
        }
        public static bool GetParameterValue(string[] Args, string ParameterName, out string Value)
        {
            List<string> argList = new List<string>(Args);

            // Case insensitive match.
            var argPosition = argList.FindIndex(
                item => item.Equals(ParameterName, StringComparison.InvariantCultureIgnoreCase)
            );

            if (argPosition == -1)
            {
                argPosition = argList.FindIndex(
                    item => item.Equals(ParameterName.Trim('-'), StringComparison.InvariantCultureIgnoreCase)
                );
            }

            if (argPosition == -1)
            {
                Value = string.Empty;
                return false;
            }

            if ((argPosition + 1) > (argList.Count - 1))
            {
                Log.Write($"The {ParameterName.Trim('-')} parameter must be followed by an argument.", IsError: true);
                throw new Invocation.MissingArgumentException($"The {ParameterName.Trim('-')} parameter must be followed by an argument.");
            }

            Value = argList[argPosition + 1];
            return true;
        }

        public static bool GetFlag(string FlagName)
        {
            return GetFlag(Invocation.Args, FlagName);
        }
        public static bool GetFlag(string[] Args, string FlagName)
        {
            List<string> argList = new List<string>(Args);

            // Case insensitive match.
            var argPosition = argList.FindIndex(
                item => item.Equals(FlagName, StringComparison.InvariantCultureIgnoreCase)
            );

            if (argPosition == -1)
                return false;
            else
                return true;
        }
    }
}