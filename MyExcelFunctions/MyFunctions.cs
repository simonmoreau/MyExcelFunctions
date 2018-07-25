using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.IO;
using System.Text.RegularExpressions;

namespace MyExcelFunctions
{
    public static class MyFunctions
    {
        [ExcelFunction(Category = "My functions", Description = "Returns the directory information for the specified path string")]
        public static object GETDIRECTORYNAME([ExcelArgument("path", Name = "path", Description = "The path of a file or directory")] string path)
        {
            try
            {
                return Path.GetDirectoryName(path);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "My functions", Description = "Returns the directory name for the specified path string.")]
        public static object GETDIRECTORY([ExcelArgument("path", Name = "path", Description = "The path of a file or directory")] string path)
        {
            try
            {
                return Path.GetFileName(Path.GetDirectoryName(path));
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "My functions", Description = "Returns the file name and extension of the specified path string.")]
        public static object GETFILENAME([ExcelArgument("path", Name = "path", Description = "The path string from which to obtain the file name and extension")] string path)
        {
            try
            {
                return Path.GetFileName(path);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "My functions", Description = "Returns the file name of the specified path string without the extension.")]
        public static object GETFILENAMEWTEXT([ExcelArgument("path", Name = "path", Description = "The path of the file")] string path)
        {
            try
            {
                return Path.GetFileNameWithoutExtension(path);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "My functions", Description = "Split a string and return the Nth item in the resulting array", HelpTopic = "Split a string and return the Nth item in the resulting array")]
        public static object SPLITSTRING(
            [ExcelArgument("string", Name = "string", Description = "The input string")] string name,
            [ExcelArgument("separator", Name = "separator", Description = "A string that delimits the substrings in this string")] string value,
            [ExcelArgument("rank", Name = "rank", Description = "The rank of the resulting substring")] int rank)
        {
            try
            {
                string[] values = new string[1] { value };
                return name.Split(values, StringSplitOptions.None)[rank];
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "My functions", Description = "Searches the specified input string for the first occurrence of the specified regular expression", HelpTopic = "Searches the specified input string for the first occurrence of the specified regular expression")]
        public static object REGEX(
            [ExcelArgument("input", Name = "input", Description = "The string to search for a match.")] string input,
            [ExcelArgument("pattern", Name = "pattern", Description = "The regular expression pattern to match.")] string pattern)
        {
            try
            {
                Match match = new Regex(pattern).Match(input);
                if (match.Success)
                {
                    return match.Value;
                }
                else
                {
                    return ExcelDna.Integration.ExcelError.ExcelErrorNA;
                }
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }


    }
}
