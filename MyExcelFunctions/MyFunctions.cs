using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.IO;

namespace MyExcelFunctions
{
    public static class MyFunctions
    {
        [ExcelFunction(Category="My functions",Description = "Returns the directory information for the specified path string")]
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

        [ExcelFunction(Category="My functions",Description = "Returns the directory name for the specified path string.")]
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

        [ExcelFunction(Category="My functions",Description = "Returns the file name and extension of the specified path string.")]
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

        [ExcelFunction(Category="My functions",Description = "Returns the file name of the specified path string without the extension.")]
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

        [ExcelFunction(Category="My functions",Description = "Split a string and return the Nth item in the resulting array",HelpTopic = "Split a string and return the Nth item in the resulting array")]
        public static object SPLITSTRING(
            [ExcelArgument("string", Name = "string", Description ="The input string")] string name,
            [ExcelArgument("separator", Name = "separator", Description = "A string that delimits the substrings in this string")] string value,
            [ExcelArgument("rank", Name = "rank", Description = "The rank of the resulting substring")] int rank)
        {
            try
            {
                string[] values = new string[1] { value };
                return name.Split(values,StringSplitOptions.None)[rank];
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        //[ExcelFunction(Category="My functions",Description="Create an image in the cell")]
        //public static object INSERTIMAGE(string path)
        //{
        //    try
        //    {
        //        Image image = new Image();
        //        image.Path = path;
        //        return image;
        //        //return name.Split(values,StringSplitOptions.None)[rank];
        //    }
        //    catch
        //    {
        //        return ExcelDna.Integration.ExcelError.ExcelErrorNA;
        //    }
        //}
    }
}
