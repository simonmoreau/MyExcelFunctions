using ExcelDna.Integration;
using ExcelFunctions.Services;
using Microsoft.Extensions.FileSystemGlobbing.Internal;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
using System.Diagnostics;

namespace ExcelFunctions
{
    public static class FileProperties
    {
        [ExcelFunction(Category = "FileProperties", Description = "Returns a property value using the given canonical name of the property.")]
        public static object GETFILEPROPERTY(
            [ExcelArgument("path", Name = "path", Description = "The path of a file")] string path,
            [ExcelArgument("canonicalName", Name = "canonicalName", Description = "The canonical name of the property")] string canonicalName)
        {
            try
            {
                ShellFile file = ShellFile.FromFilePath(path);
                IShellProperty property = file.Properties.GetProperty(canonicalName);

                return property.ValueAsObject;
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "FileProperties", Description = "Gets the collection of all the default properties for this file.")]
        public static object GETFILEAVAILABLEPROPERTIES(
    [ExcelArgument("path", Name = "path", Description = "The path of a file")] string path)
        {
            try
            {
                ShellFile file = ShellFile.FromFilePath(path);

                // Read and Write:
                List<IShellProperty> properties = file.Properties.DefaultPropertyCollection.ToList();

                string[] propertiesName = properties.Select(p => p.CanonicalName).ToArray();

                if (propertiesName.Length == 0) return ExcelDna.Integration.ExcelError.ExcelErrorNull;

                object[,] outputTable = new object[propertiesName.Length, 1];

                int l = 0;
                foreach (object obj in propertiesName)
                {
                    outputTable[l, 0] = obj;
                    l++;
                }
                return outputTable;
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "FileProperties", Description = "Returns a property value using the given canonical name of the property.")]
        public static object SETFILEPROPERTY(
    [ExcelArgument("path", Name = "path", Description = "The path of a file")] string path,
    [ExcelArgument("canonicalName", Name = "canonicalName", Description = "The canonical name of the property")] string canonicalName,
    [ExcelArgument("value", Name = "value", Description = "The value associated with the property.")] object value)
        {
            try
            {
                ShellFile file = ShellFile.FromFilePath(path);
                IShellProperty property = file.Properties.GetProperty(canonicalName);

                ShellPropertyWriter propertyWriter = file.Properties.GetPropertyWriter();
                propertyWriter.WriteProperty(property.PropertyKey, value);
                propertyWriter.Close();

                return "Value updated";
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }
    }
}
