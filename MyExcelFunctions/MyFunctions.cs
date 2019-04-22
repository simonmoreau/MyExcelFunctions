using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;


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

        [ExcelFunction(Category = "My functions", Description = "Returns the extension of the specified path string.")]
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

        [ExcelFunction(Category = "My functions", Description = "Converts the string representation of a date and time to its System.DateTime equivalent.", HelpTopic = "Converts the string representation of a date and time to its System.DateTime equivalent.")]
        public static object PARSEDATE(
            [ExcelArgument("input", Name = "input", Description = "A string that contains a date and time to convert.")] string input,
            [ExcelArgument("format", Name = "format", Description = "[optional]A format specifier that defines the required format of input.")] string format)
        {
            try
            {
                if (format == "")
                {
                    return DateTime.Parse(input);
                }
                else
                {
                    CultureInfo provider = CultureInfo.InvariantCulture;
                    return DateTime.ParseExact(input, format, provider, DateTimeStyles.None);
                }
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "My functions", Description = "Insert a picture.", HelpTopic = "Insert a picture.")]
        public static object INSERTPICTURE(
    [ExcelArgument("path", Name = "path", Description = "A path to the picture.")] string path,
    [ExcelArgument("width", Name = "width", Description = "[optional] The width of the picture.")] string width,
    [ExcelArgument("height", Name = "height", Description = "[optional] The height of the picture.")] string height)
        {
            try
            {
                
                Excel.Application excelApplication = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Worksheet ws = (Excel.Worksheet)excelApplication.ActiveSheet;

                ExcelReference caller = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);

                Excel.Range oRange = (Excel.Range)ws.Cells[caller.RowFirst + 1, caller.ColumnFirst + 1];

                try
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("fr-FR");
                    //Get range value
                    int? value = null;

                    //System.Globalization.CultureInfo cultureInfo = new System.Globalization.CultureInfo("en-US");
                    //var test = oRange.GetType().InvokeMember("Value2", BindingFlags.InvokeMethod, null, oRange, null);

                    //var cellTest = ws.Cells[caller.RowFirst + 1, caller.ColumnFirst + 1].Values2;
                    var cellValue = (string)oRange.Value2;
                    value = Convert.ToInt16(cellValue);

                    if (value != null)
                    {
                        foreach (Excel.Shape shape in ws.Shapes)
                        {
                            if (shape.ID == oRange.Value)
                            {
                                shape.Delete();
                            }
                        }
                    }
                }
                catch
                {

                }



                float Left = (float)((double)oRange.Left);
                float Top = (float)((double)oRange.Top);

                System.Drawing.Image img = System.Drawing.Image.FromFile(path);

                double ratio =  img.Width / img.Height;
                float pictureHeight = 50;
                float pictureWidth = (float)Math.Round(pictureHeight * ratio);

                if (height == "" && width == "")
                {

                }
                else if (height == "")
                {
                    pictureWidth = (float)Convert.ToInt16(width);
                    pictureHeight = (float)Math.Round(pictureWidth / ratio);
                }
                else if (width == "")
                {
                    
                    pictureHeight = (float)Convert.ToInt16(height);
                    pictureWidth = (float)Math.Round(pictureHeight * ratio);
                }

                

                Excel.Shape picture = ws.Shapes.AddPicture(path, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, pictureWidth, pictureHeight);
                
                return picture.ID.ToString(); ;
            }
            catch (Exception ex)
            {
                string test = ex.Message;
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "My functions", Description = "Converts a given number into its full text in french.", HelpTopic = "Converts a given number into its full text in french.")]
        public static object CONVERTTOFRENCH(
    [ExcelArgument("input", Name = "input", Description = "The number to be spelled")] int input)
        {
            try
            {
                TexteEnLettre texteEnLettre = new TexteEnLettre();
                string texte = texteEnLettre.IntToFr(input);
                //Clean text
                texte = texte.Replace("  ", " ");
                texte = texte.Trim(' ');
                // Ajoute une majuscule au début
                return texte.First().ToString().ToUpper() + texte.Substring(1);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }


        [ExcelFunction(Category = "My functions", Description = "Create a table from sets of possible values.", HelpTopic = "Create a table from sets of possible values.")]
        public static object GENERATELISTFROMVALUES(
    [ExcelArgument("values", Name = "values", Description = "A series of values in columns.")] object values)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;
                    //Create a list of list of string
                    List<List<object>> columns = new List<List<object>>();

                    for (int i = 0; i < inputArray.GetLength(1); i++)
                    {
                        List<object> column = new List<object>();

                        for (int j = 0; j < inputArray.GetLength(0); j++)
                        {
                            if (inputArray[j, i] != ExcelDna.Integration.ExcelEmpty.Value)
                            {
                                column.Add(inputArray[j, i]);
                            }
                        }
                        columns.Add(column);
                    }

                    int lineNumber = columns.Aggregate(1, (x, y) => x * y.Count);
                    int columnNumber = columns.Count;

                    object[,] outputTable = new object[lineNumber, columnNumber];

                    //for (int i = 0; i < outputTable.GetLength(0); i++)
                    //{
                    //    for (int j = 0; j < outputTable.GetLength(1); j++)
                    //    {
                    //        if (inputArray[i, j] != ExcelDna.Integration.ExcelEmpty.Value)
                    //        {
                    //            outputTable[i, j] = 2;//columns[i][j];
                    //        }
                    //    }
                    //}

                    //List<List<object>> query = CrossJoinList(columns[0], columns[1]);

                    IEnumerable<IEnumerable<object>> query = CartesianProduct(columns);


                    int k = 0;
                    foreach (IEnumerable<object> objects in query)
                    {
                        int l = 0;
                        foreach (object obj in objects)
                        {
                            outputTable[k, l] = obj;
                            l++;
                        }
                        k++;
                    }
                    return outputTable;
                }
                catch
                {
                    return new object[,] { { ExcelDna.Integration.ExcelError.ExcelErrorNA } };
                }
            }
            else
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }

        }

        private static List<List<object>> CrossJoinList(List<object> list1, List<object> list2)
        {
            List<List<object>> query = list1.SelectMany(list0 => list2, (object0, object1) => new List<object>() { object0, object1 }).ToList<List<object>>();

            return query;
        }

        private static IEnumerable<IEnumerable<T>> CartesianProduct<T>(this IEnumerable<IEnumerable<T>> sequences)
        {
            if (sequences == null)
            {
                return null;
            }

            IEnumerable<IEnumerable<T>> emptyProduct = new[] { Enumerable.Empty<T>() };

            return sequences.Aggregate(
                emptyProduct,
                (accumulator, sequence) => accumulator.SelectMany(
                    accseq => sequence,
                    (accseq, item) => accseq.Concat(new[] { item })));
        }

    }
}
