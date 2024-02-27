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
                return System.IO.File.GetCreationTime(path);
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
                return System.IO.File.GetLastWriteTime(path);
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

        [ExcelFunction(Category = "IO", Description = "Gets the size of an image WidthxHeight.")]
        public static object IMAGESIZE([ExcelArgument("path", Name = "path", Description = "The path to the image file.")] string path)
        {
            try
            {
                System.Drawing.Image image = System.Drawing.Image.FromStream(File.OpenRead(path), false, false);

                if (image == null)
                {
                    return ExcelDna.Integration.ExcelError.ExcelErrorNull;
                }

                return $"{image.Width}x{image.Height}";
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

        [ExcelFunction(Category = "IO", Description = "Replace invalid caracters in a file name.")]
        public static object GETVALIDFILENAME(
    [ExcelArgument("filename", Name = "filename", Description = "The filename to modify.")] string filename)
        {
            try
            {
                foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                {
                    filename = filename.Replace(c, '_');
                }

                return filename;
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
                if (System.IO.File.Exists(destFileName))
                {
                    if (overrideDestination)
                    {
                        System.IO.File.Delete(destFileName);
                    }
                    else
                    {
                        return sourceFileName;
                    }
                }

                // Move the file.
                System.IO.File.Move(sourceFileName, destFileName);
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
                System.IO.File.Copy(sourceFileName, destFileName, overrideDestination);
                return destFileName;
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "IO", Description = "Opens a text file, reads all lines of the file, and then closes the file.")]
        public static object READALLTEXT([ExcelArgument("path", Name = "path", Description = "The file to open for reading.")] string path)
        {
            try
            {
                return System.IO.File.ReadAllText(path);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "IO", Description = "Opens a text file, reads all lines of the file, and then closes the file.")]
        public static object READALLLINES([ExcelArgument("path", Name = "path", Description = "The file to open for reading.")] string path)
        {
            try
            {
                string[] lines = System.IO.File.ReadAllLines(path);

                object[,] outputTable = new object[lines.Length, 1];

                int l = 0;
                foreach (object obj in lines)
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

        [ExcelFunction(Category = "IO", Description = "Creates a new file, writes the specified string to the file, and then closes the file. If the target file already exists, it is overwritten.")]
        public static object WRITEALLTEXT(
            [ExcelArgument("path", Name = "path", Description = "The file to write to.")] string path,
            [ExcelArgument("contents", Name = "contents", Description = "The string to write to the file.")] string contents)
        {
            try
            {
                System.IO.File.WriteAllText(path, contents);
                return $"Text written to {path}";
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "IO", Description = "Creates all directories and subdirectories in the specified path unless they already exist.")]
        public static object CREATEDIRECTORY(
    [ExcelArgument("path", Name = "path", Description = "The directory to create..")] string path)
        {
            try
            {
                Directory.CreateDirectory(path);
                return $"{path}";
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "IO", Description = "Creates all directories and subdirectories in the specified path unless they already exist.")]
        public static object COPYDIRECTORY(
[ExcelArgument("sourceDirectoryName", Name = "sourceDirectoryName", Description = "The directory to be copied.")] string sourceDirectoryName,
[ExcelArgument("destinationDirectoryName", Name = "destinationDirectoryName", Description = "The location to which the directory contents should be copied.")] string destinationDirectoryName,
[ExcelArgument("[recursive]", Name = "[recursive]", Description = "true to include subfolders.")] object recursive)
        {
            try
            {
                bool recursiveBool = Optional.Check(recursive, false);
                CopyDirectory(sourceDirectoryName, destinationDirectoryName, recursiveBool);
                return destinationDirectoryName;
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        private static void CopyDirectory(string sourceDir, string destinationDir, bool recursive)
        {
            // Get information about the source directory
            DirectoryInfo directoryInfo = new DirectoryInfo(sourceDir);

            // Check if the source directory exists
            if (!directoryInfo.Exists)
                throw new DirectoryNotFoundException($"Source directory not found: {directoryInfo.FullName}");

            // Cache directories before we start copying
            DirectoryInfo[] dirs = directoryInfo.GetDirectories();

            // Create the destination directory
            Directory.CreateDirectory(destinationDir);

            // Get the files in the source directory and copy to the destination directory
            foreach (FileInfo file in directoryInfo.GetFiles())
            {
                string targetFilePath = Path.Combine(destinationDir, file.Name);
                file.CopyTo(targetFilePath,true);
            }

            // If recursive and copying subdirectories, recursively call this method
            if (recursive)
            {
                foreach (DirectoryInfo subDir in dirs)
                {
                    string newDestinationDir = Path.Combine(destinationDir, subDir.Name);
                    CopyDirectory(subDir.FullName, newDestinationDir, true);
                }
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
