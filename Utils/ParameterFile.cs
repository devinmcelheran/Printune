using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace Printune
{
    public class ParameterFile
    {
        private Dictionary<string, string> _parameters;

        public ParameterFile()
        {
            _parameters = new Dictionary<string, string>();
        }

        public ParameterFile(string filePath)
        {
            try
            {
                var content = File.ReadAllText(filePath);
                _parameters = JsonConvert.DeserializeObject<Dictionary<string, string>>(content);
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

        public string GetParameter(string parameterName)
        {
            if (string.IsNullOrEmpty(parameterName))
                throw new ArgumentException("Parameter name cannot be null or empty.", nameof(parameterName));

            string value;
            if (_parameters.TryGetValue(parameterName, out value))
            {
                return value;
            }
            else
            {
                throw new ArgumentException($"Parameter '{parameterName}' not found in the parameter file.");
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