using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcelFunctions
{
    public static class IO
    {
        [ExcelFunction(Category = "IO", Description = "Returns the directory information for the specified path string")]
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

        [ExcelFunction(Category = "IO", Description = "Returns the directory name for the specified path string.")]
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

        [ExcelFunction(Category = "IO", Description = "Returns the file name and extension of the specified path string.")]
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

        [ExcelFunction(Category = "IO", Description = "Returns the file name of the specified path string without the extension.")]
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

        [ExcelFunction(Category = "IO", Description = "Returns the extension of the specified path string.")]
        public static object GETEXTENSION([ExcelArgument("path", Name = "path", Description = "The path string from which to get the extension.")] string path)
        {
            try
            {
                return Path.GetExtension(path);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "IO", Description = "Changes the extension of a path string.")]
        public static object CHANGEEXTENSION(
            [ExcelArgument("path", Name = "path", Description = "The path information to modify. The path cannot contain any of the characters defined in System.IO.Path.GetInvalidPathChars.")] string path,
            [ExcelArgument("extension", Name = "extension", Description = "The new extension (with or without a leading period). Specify null to remove an existing extension from path.")] string extension)
        {
            try
            {
                return Path.ChangeExtension(path, extension);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "IO", Description = "Returns the names of files (including their paths) in the specified directory.")]
        public static object GETFILES(
            [ExcelArgument("path", Name = "path", Description = "The relative or absolute path to the directory to search. This string is not case-sensitive.")] string path,
            [ExcelArgument("[searchPattern]", Name = "[searchPattern]", Description = "The search string to match against the names of files in path. This parameter can contain a combination of valid literal path and wildcard (* and ?) characters,but it doesn't support regular expressions.")] object searchPattern,
            [ExcelArgument("[allDirectories]", Name = "[allDirectories]", Description = "The search string to match against the names of files in path. This parameter can contain a combination of valid literal path and wildcard (* and ?) characters,but it doesn't support regular expressions.")] object allDirectories)
        {
            try
            {
                
                string pattern = Optional.Check(searchPattern, "");
                if (pattern == "") pattern = "*";

                bool allDirectoriesBool = Optional.Check(allDirectories, false);
                SearchOption searchOption = SearchOption.TopDirectoryOnly;
                if (allDirectoriesBool) { searchOption= SearchOption.AllDirectories; }

                string[] paths = Directory.GetFiles(path, pattern, searchOption);


                if (paths.Length== 0) return ExcelDna.Integration.ExcelError.ExcelErrorNull;

                object[,] outputTable = new object[paths.Length, 1];

                int l = 0;
                foreach (object obj in paths)
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

        [ExcelFunction(Category = "IO", Description = "Returns the number of files in the specified directory.")]
        public static object COUNTFILES(
    [ExcelArgument("path", Name = "path", Description = "The relative or absolute path to the directory to search. This string is not case-sensitive.")] string path,
    [ExcelArgument("[searchPattern]", Name = "[searchPattern]", Description = "The search string to match against the names of files in path. This parameter can contain a combination of valid literal path and wildcard (* and ?) characters,but it doesn't support regular expressions.")] object searchPattern,
              [ExcelArgument("[allDirectories]", Name = "[allDirectories]", Description = "The search string to match against the names of files in path. This parameter can contain a combination of valid literal path and wildcard (* and ?) characters,but it doesn't support regular expressions.")] object allDirectories)
        {
            try
            {
                string pattern = Optional.Check(searchPattern, "");
                if (pattern == "") pattern = "*";

                bool allDirectoriesBool = Optional.Check(allDirectories, false);
                SearchOption searchOption = SearchOption.TopDirectoryOnly;
                if (allDirectoriesBool) { searchOption = SearchOption.AllDirectories; }

                string[] paths = Directory.GetFiles(path, pattern, searchOption);

                return paths.Length;

            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

    }
}

//     The search string to match against the names of files in path. This parameter can contain a combination of valid literal path and wildcard (* and ?) characters,but it doesn't support regular expressions.

//
// Summary:
//     Returns the names of files (including their paths) in the specified directory.
//
// Parameters:
//   path:
//     
//     
//
// Returns:
//     An array of the full names (including paths) for the files in the specified directory,
//     or an empty array if no files are found.
//
// Exceptions:
//   T:System.IO.IOException:
//     path is a file name. -or- A network error has occurred.
//
//   T:System.UnauthorizedAccessException:
//     The caller does not have the required permission.
//
//   T:System.ArgumentException:
//     path is a zero-length string, contains only white space, or contains one or more
//     invalid characters. You can query for invalid characters by using the System.IO.Path.GetInvalidPathChars
//     method.
//
//   T:System.ArgumentNullException:
//     path is null.
//
//   T:System.IO.PathTooLongException:
//     The specified path, file name, or both exceed the system-defined maximum length.
//
//   T:System.IO.DirectoryNotFoundException:
//     The specified path is not found or is invalid (for example, it is on an unmapped
//     drive).