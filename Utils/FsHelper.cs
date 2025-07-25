using System;
using System.IO;

namespace Printune
{
    public static class FsHelper
    {
        public static string CreateDirectory(string DirectoryPath)
        {
            if (DirectoryPath == null)
                throw new NullReferenceException("A null value was passed instead of a valid path.");

            if (Directory.Exists(DirectoryPath))
                return new DirectoryInfo(DirectoryPath).FullName;

            if (!Directory.Exists(Path.GetDirectoryName(DirectoryPath)))
                CreateDirectory(Path.GetDirectoryName(DirectoryPath));

            return Directory.CreateDirectory(DirectoryPath).FullName;
        }
    }
}