using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFunctions.Services
{
    internal static class Optional
    {
        internal static string Check(object arg, string defaultValue)
        {
            if (arg is string)
                return (string)arg;
            else if (arg.GetType().Name == typeof(ExcelMissing).Name)
                return defaultValue;
            else
                return arg.ToString();  // Or whatever you want to do here....

            // Perhaps check for other types and do whatever you think is right ....
            //else if (arg is double)
            //    return "Double: " + (double)arg;
            //else if (arg is bool)
            //    return "Boolean: " + (bool)arg;
            //else if (arg is ExcelError)
            //    return "ExcelError: " + arg.ToString();
            //else if (arg is object[,](,))
            //    // The object array returned here may contain a mixture of types,
            //    // reflecting the different cell contents.
            //    return string.Format("Array[{0},{1}]({0},{1})",
            //      ((object[,](,)(,))arg).GetLength(0), ((object[,](,)(,))arg).GetLength(1));
            //else if (arg is ExcelEmpty)
            //    return "<<Empty>>"; // Would have been null
            //else if (arg is ExcelReference)
            //  // Calling xlfRefText here requires IsMacroType=true for this function.
            //                return "Reference: " +
            //                     XlCall.Excel(XlCall.xlfReftext, arg, true);
            //            else
            //                return "!? Unheard Of ?!";
        }

        internal static double Check(object arg, double defaultValue)
        {
            if (arg is double)
                return (double)arg;
            else if (arg is ExcelMissing)
                return defaultValue;
            else
                throw new ArgumentException();  // Will return #VALUE to Excel

        }

        internal static int Check(object arg, int defaultValue)
        {
            if (arg is int)
                return Convert.ToInt16(arg);
            else if (arg.GetType().Name == typeof(ExcelMissing).Name )
                return defaultValue;
            else
                throw new ArgumentException();  // Will return #VALUE to Excel
        }

        internal static bool Check(object arg, bool defaultValue)
        {
            if (arg is bool)
                return (bool)arg;
            else if (arg is ExcelMissing)
                return defaultValue;
            else
                throw new ArgumentException();  // Will return #VALUE to Excel

        }

        internal static string[]? Check(object arg)
        {
            if (arg is object[,])
            {
                List<string> list = new List<string>();
                object[,] argArray = (object[,])arg;

                foreach (string value in argArray)
                {
                    list.Add((string)value);
                }
                return list.ToArray();
            }
            else if (arg is ExcelMissing)
                return null;
            else
                throw new ArgumentException();  // Or defaultValue or whatever
        }

        // This one is more tricky - we have to do the double->Date conversions ourselves
        internal static DateTime Check(object arg, DateTime defaultValue)
        {
            if (arg is double)
                return DateTime.FromOADate((double)arg);    // Here is the conversion
            else if (arg is string)
                return DateTime.Parse((string)arg);
            else if (arg is ExcelMissing)
                return defaultValue;

            else
                throw new ArgumentException();  // Or defaultValue or whatever
        }
    }
}
