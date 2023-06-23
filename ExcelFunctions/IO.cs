using ExcelDna.Integration;
using ExcelFunctions.Services;
using System.Diagnostics;

namespace ExcelFunctions
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

        [ExcelFunction(Category = "IO", Description = "Retrieves the parent directory of the specified path, including both absolute and relative paths.")]
        public static object GETPARENTDIRECTORY([ExcelArgument("path", Name = "path", Description = "The path for which to retrieve the parent directory.")] string path)
        {
            try
            {
                DirectoryInfo directoryInfo = Directory.GetParent(path);
                if (directoryInfo == null) return "";

                return directoryInfo.FullName;
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

        [ExcelFunction(Category = "IO", Description = "Determines whether the specified file exists.")]
        public static object FILEEXISTS([ExcelArgument("path", Name = "path", Description = "The file to check.")] string path)
        {
            try
            {
                return File.Exists(path);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "IO", Description = "Returns the creation time of the specified file or directory.")]
        public static object FILEGETCREATIONTIME([ExcelArgument("path", Name = "path", Description = "The file or directory for which to obtain creation date and time information.")] string path)
        {
            try
            {
                return File.GetCreationTime(path);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "IO", Description = "Returns the last write date and time of the specified file or directory.")]
        public static object FILEGETLASTWRITETIME([ExcelArgument("path", Name = "path", Description = "The file or directory for which to obtain write date and time information.")] string path)
        {
            try
            {
                return File.GetLastWriteTime(path);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "IO", Description = "Gets the size, in bytes, of the current file.")]
        public static object FILELENGHT([ExcelArgument("path", Name = "path", Description = "The file for which to obtain its size.")] string path)
        {
            try
            {
                return new System.IO.FileInfo(path).Length;
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "IO", Description = "Gets a human-readable file size")]
        public static object PARSEBYSIZE([ExcelArgument("size", Name = "size", Description = "The input size, in bytes.")] double size)
        {
            try
            {
                string[] sizes = { "B", "KB", "MB", "GB", "TB" };
                int order = 0;
                while (size >= 1024 && order < sizes.Length - 1)
                {
                    order++;
                    size = size / 1024;
                }

                // Adjust the format string to your preferences. For example "{0:0.#}{1}" would
                // show a single decimal place, and no space.
                string result = String.Format("{0:0.##} {1}", size, sizes[order]);

                return result;
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

        [ExcelFunction(Category = "IO", Description = "Moves a specified file to a new location, providing the options to specify a new file name and to overwrite the destination file if it already exists.")]
        public static object MOVEFILE(
[ExcelArgument("sourceFileName", Name = "sourceFileName", Description = "The name of the file to move. Can include a relative or absolute path.")] string sourceFileName,
[ExcelArgument("destFileName", Name = "destFileName", Description = "The new path and name for the file.")] string destFileName,
      [ExcelArgument("[override]", Name = "[override]", Description = "Erase the target if it already exist.")] object overrideDest)
        {
            try
            {
                bool overrideDestination = Optional.Check(overrideDest, false);
                // Ensure that the target does not exist.
                if (File.Exists(destFileName))
                {
                    if (overrideDestination)
                    {
                        File.Delete(destFileName);
                    }
                    else
                    {
                        return sourceFileName;
                    }
                }

                // Move the file.
                File.Move(sourceFileName, destFileName);
                return destFileName;
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "IO", Description = "Copies an existing file to a new file. Overwriting a file of the same name is allowed.")]
        public static object COPYFILE(
[ExcelArgument("sourceFileName", Name = "sourceFileName", Description = "The file to copy.")] string sourceFileName,
[ExcelArgument("destFileName", Name = "destFileName", Description = "The name of the destination file. This cannot be a directory.")] string destFileName,
      [ExcelArgument("[override]", Name = "[override]", Description = "true if the destination file can be overwritten; otherwise, false.")] object overrideDest)
        {
            try
            {
                bool overrideDestination = Optional.Check(overrideDest, false);

                // Move the file.
                File.Copy(sourceFileName, destFileName, overrideDestination);
                return destFileName;
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        public static object SleepAsync(string ms)
        {
            return ExcelAsyncUtil.Run("SleepAsync", ms, delegate
            {
                Debug.Print("{1:HH:mm:ss.fff} Sleeping for {0} ms", ms, DateTime.Now);
                Thread.Sleep(int.Parse(ms));

                Debug.Print("{1:HH:mm:ss.fff} Done sleeping {0} ms", ms, DateTime.Now);
                return "Woke Up at " + DateTime.Now.ToString("1:HH:mm:ss.fff");
            });

        }


    }
}
