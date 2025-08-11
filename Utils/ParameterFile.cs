using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

/// <summary>
/// This class represents a parameter file that can be used to store key-value pairs.
/// The paramter file can be provided instead of arguments at the command line.
/// </summary>
namespace Printune
{
    public class ParameterFile
    {
        private Dictionary<string, object> _parameters = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        public ParameterFile() { }

        public ParameterFile(string filePath)
        {
            try
            {
                var content = File.ReadAllText(filePath);
                var fileParams = JsonConvert.DeserializeObject<Dictionary<string, object>>(content);

                foreach (var param in fileParams)
                {
                    _parameters[param.Key] = param.Value;
                }
            }
            catch (Exception ex)
            {
                throw new JsonException($"Failed to read or parse the parameter file at {filePath}.", ex);
            }
        }

        public void AddParameter(string parameterName, string parameterValue)
        {
            if (string.IsNullOrEmpty(parameterName))
                throw new ArgumentException("Parameter name cannot be null or empty.", nameof(parameterName));

            _parameters[parameterName] = parameterValue;
        }

        public object GetParameter(string parameterName)
        {
            if (string.IsNullOrEmpty(parameterName))
                throw new ArgumentException("Parameter name cannot be null or empty.", nameof(parameterName));

            object value;
            if (_parameters.TryGetValue(parameterName, out value))
            {
                return value;
            }
            else
            {
                throw new KeyNotFoundException($"Parameter '{parameterName}' not found in the parameter file.");
            }
        }

        public void WriteToFile(string filePath)
        {
            try
            {
                var content = JsonConvert.SerializeObject(_parameters, Formatting.Indented);
                File.WriteAllText(filePath, content);
            }
            catch (Exception ex)
            {
                throw new IOException($"Failed to write content to parameter file {filePath}", ex);
            }
        }
    }
}