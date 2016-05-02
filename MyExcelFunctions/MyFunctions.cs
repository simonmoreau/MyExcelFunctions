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
        [ExcelFunction(Category="My functions",Description = "My first .NET function")]
        public static string SayHello(string name)
        {
            return "Hello 2 " + name;
        }

        [ExcelFunction(Category="My functions",Description = "Returns the directory information for the specified path string.")]
        public static object GETDIRECTORYNAME(string name)
        {
            try
            {
                return Path.GetDirectoryName(name);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category="My functions",Description = "Returns the directory name for the specified path string.")]
        public static object GETDIRECTORY(string name)
        {
            try
            {
                return Path.GetFileName(Path.GetDirectoryName(name));
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category="My functions",Description = "Returns the file name and extension of the specified path string.")]
        public static object GETFILENAME(string name)
        {
            try
            {
                return Path.GetFileName(name);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category="My functions",Description = "Returns the file name of the specified path string without the extension.")]
        public static object GETFILENAMEWTEXT(string name)
        {
            try
            {
                return Path.GetFileNameWithoutExtension(name);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category="My functions",Description = "Split a string and return the Nth item in the resulting array.")]
        public static object SPLITSTRING(
            [ExcelArgument(Name = "string")] string name,
            [ExcelArgument(Name = "A string that delimit the substring")] string value,
            [ExcelArgument(Name = "The rank of the resulting substring")] int rank)
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
