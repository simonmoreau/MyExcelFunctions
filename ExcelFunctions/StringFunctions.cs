using ExcelDna.Integration;
using ExcelFunctions.Services;
using FuzzySharp;
using Microsoft.Extensions.FileSystemGlobbing.Internal;
using System;
using System.Data.Common;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ExcelFunctions
{
    public static class StringFunctions
    {
        [ExcelFunction(Category = "String", Description = "Split a string and return the Nth item in the resulting array", HelpTopic = "Split a string and return the Nth item in the resulting array")]
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

        [ExcelFunction(Category = "String", Description = "Splits a string into substrings based on a specified delimiting string.", HelpTopic = "Splits a string into substrings based on a specified delimiting string.")]
        public static object SPLITVALUE(
[ExcelArgument("string", Name = "string", Description = "The input string")] string input,
[ExcelArgument("separator", Name = "separator", Description = "A string that delimits the substrings in this string.")] string separator)
        {
            try
            {
                string[] values = input.Split(separator, StringSplitOptions.None);

                object[,] outputTable = new object[values.Length, 1];

                int l = 0;
                foreach (string value in values)
                {
                    outputTable[l, 0] = value;
                    l++;
                }

                return outputTable;
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "String", Description = "Converts the string representation of a date and time to its System.DateTime equivalent.", HelpTopic = "Converts the string representation of a date and time to its System.DateTime equivalent.")]
        public static object PARSEDATE(
            [ExcelArgument("input", Name = "input", Description = "A string that contains a date and time to convert.")] string input,
            [ExcelArgument("[format]", Name = "[format]", Description = "[optional]A format specifier that defines the required format of input.")] string format)
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

        [ExcelFunction(Category = "String", Description = "Returns a value indicating whether a specified substring occurs within this string.", HelpTopic = "Returns a value indicating whether a specified substring occurs within this string.")]
        public static object CONTAINS(
    [ExcelArgument("input", Name = "input", Description = "The input string.")] string input,
    [ExcelArgument("value", Name = "value", Description = "The string to seek.")] string value)
        {
            try
            {
                return input.Contains(value);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "String", Description = "Escape non ASCII characters with their unicode value", HelpTopic = "Escape non ASCII characters with their unicode value")]
        public static object ENCODENONASCIICHARACTERS(
    [ExcelArgument("input", Name = "input", Description = "The string to escape.")] string input)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                foreach (char c in input)
                {
                    if (c > 127)
                    {
                        // This character is too big for ASCII
                        string encodedValue = "\\u" + ((int)c).ToString("x4");
                        sb.Append(encodedValue);
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
                return sb.ToString();
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "String", Description = "Replace string in text based on an array.", HelpTopic = "Replace string in text based on an array.")]
        public static object FINDANDREPLACE(
[ExcelArgument("table", Name = "table", Description = "The table where to look for values")] object values,
[ExcelArgument("column", Name = "column", Description = "The column where to find replacement values")] object column,
[ExcelArgument("text", Name = "text", Description = "The text to be replaced")] object text)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;
                    string textValue = (string)text;
                    int columnIndex = 0;
                    bool parseResult = int.TryParse(column.ToString(), out columnIndex);

                    if (!parseResult) return ExcelDna.Integration.ExcelError.ExcelErrorNA;

                    Dictionary<string, string> keyDictionary = new Dictionary<string, string>();

                    for (int i = 0; i < inputArray.GetLength(0); i++)
                    {
                        string key = inputArray[i, 0].ToString();
                        if (textValue.Contains(key))
                        {
                            string replacementValue = inputArray[i, columnIndex].ToString();
                            keyDictionary.Add(key, replacementValue);
                        }
                    }

                    string textReturn = textValue;
                    foreach (string key in keyDictionary.Keys)
                    {
                        textReturn = textReturn.Replace(key, keyDictionary[key]);
                    }

                    return textReturn;

                }
                catch (Exception ex)
                {
                    return new object[,] { { ExcelDna.Integration.ExcelError.ExcelErrorNA } };
                }
            }
            else
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }

        }

        [ExcelFunction(Category = "String", Description = "Replace string in text based on an array.", HelpTopic = "Replace string in text based on an array.")]
        public static object FUZZYLOOKUP(
[ExcelArgument("choices", Name = "choices", Description = "An array of string to look into")] object choices,
[ExcelArgument("text", Name = "text", Description = "The text to search")] string text,
[ExcelArgument("[threshold]", Name = "[threshold]", Description = "The minimun score to get a match.")] object threshold)
        {
            if (choices is object[,])
            {
                try
                {
                    int thresholdValue = Optional.Check(threshold, 1);

                    object[,] inputArray = (object[,])choices;
                    string?[] inputColumn = Enumerable.Range(0, inputArray.GetLength(0)).Select(x => Convert.ToString(inputArray[x, 0])).ToArray();

                    FuzzySharp.Extractor.ExtractedResult<string> extractedResult = Process.ExtractOne(text, inputColumn);

                    if (extractedResult.Score > thresholdValue)
                    {
                        return extractedResult.Value;
                    }
                    else
                    {
                        return new object[,] { { ExcelDna.Integration.ExcelError.ExcelErrorNA } };
                    }
                }
                catch (Exception ex)
                {
                    return new object[,] { { ExcelDna.Integration.ExcelError.ExcelErrorNA } };
                }
            }
            else
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }

        }

        [ExcelFunction(Category = "String", Description = "Find the first occurence of a text in a table and return the content of the cell.", HelpTopic = "Find the first occurence of a text in a table and return the content of the cell.")]
        public static object FINDINLIST(
[ExcelArgument("table", Name = "table", Description = "The table where to look for values")] object values,
[ExcelArgument("text", Name = "text", Description = "The text to be found")] object text)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;
                    string textValue = (string)text;

                    for (int i = 0; i < inputArray.GetLength(0); i++)
                    {
                        for (int j = 0; j < inputArray.GetLength(1); j++)
                        {
                            string cellContent = inputArray[i, j].ToString();
                            if (cellContent.Contains(textValue))
                            {
                                return inputArray[i, j].ToString();
                            }
                        }
                    }

                    return ExcelDna.Integration.ExcelError.ExcelErrorNull;

                }
                catch (Exception ex)
                {
                    return new object[,] { { ExcelDna.Integration.ExcelError.ExcelErrorNA } };
                }
            }
            else
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }

        }

        [ExcelFunction(Category = "String", Description = "Converts a given number into its full text in french.", HelpTopic = "Converts a given number into its full text in french.")]
        public static object CONVERTTOFRENCH(
[ExcelArgument("input", Name = "input", Description = "The number to be spelled")] double value,
[ExcelArgument("portion", Name = "portion", Description = "[optional] Type 0 for the interger portion, 1 for decimal portion, 2 for both, 3 for 2 digits. Default is 0")] int portion)
        {
            try
            {
                TexteEnLettre texteEnLettre = new TexteEnLettre();
                if (portion == 3)
                {
                    value = Math.Round(value, 2);
                }

                int wholePart = (int)Math.Truncate(value);
                decimal decimalValue = (decimal)(value - Math.Truncate(value));
                long decimalPart = 0;
                long twoDigitDecimalPart = 0;
                if (decimalValue != 0)
                {
                    string decimalString = value.ToString().Split(System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator.ToCharArray())[1];
                    if (decimalString.Length > 19)
                    {
                        decimalString = decimalString.Substring(0, 19);
                    }
                    string TwoDigitdecimalString = decimalString;

                    if (decimalString.Length == 0)
                    {
                        TwoDigitdecimalString = "0";
                    }
                    else if (decimalString.Length == 1)
                    {
                        TwoDigitdecimalString = decimalString + "0";
                    }
                    else
                    {
                        TwoDigitdecimalString = decimalString.Substring(0, 2);
                    }

                    decimalPart = Convert.ToInt64(decimalString);
                    twoDigitDecimalPart = Convert.ToInt64(TwoDigitdecimalString);
                }


                string texte = texteEnLettre.IntToFr(wholePart);

                if (portion == 0)
                {
                    texte = texteEnLettre.IntToFr(wholePart);
                }
                else if (portion == 1)
                {
                    texte = texteEnLettre.IntToFr(decimalPart);
                }
                else if (portion == 2)
                {
                    texte = texteEnLettre.IntToFr(wholePart) + " virgule " + texteEnLettre.IntToFr(decimalPart);
                }
                else if (portion == 3)
                {
                    texte = texteEnLettre.IntToFr(twoDigitDecimalPart);
                }

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

        [ExcelFunction(Category = "String", Description = "Converts the specified name to CamelCase.", HelpTopic = "Converts the specified name to CamelCase.")]
        public static object CAMELCASE(
[ExcelArgument("name", Name = "name", Description = "The name to convert.")] string name)
        {
            try
            {
                return System.Text.Json.JsonNamingPolicy.CamelCase.ConvertName(name);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "String", Description = "Converts the specified name to CamelCase.", HelpTopic = "Converts the specified name to CamelCase.")]
        public static object TITLECASE(
[ExcelArgument("name", Name = "name", Description = "The name to convert.")] string name)
        {
            try
            {
                TextInfo ti = CultureInfo.CurrentCulture.TextInfo;
                return ti.ToTitleCase(name);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "String", Description = "Return a globally unique identifier (GUID).", HelpTopic = "Return a globally unique identifier (GUID).")]
        public static object GUID(
[ExcelArgument("[type]", Name = "[type]", Description = "(Optional) The type of guid, long (0) or short (1) to return.")] int type)
        {
            try
            {
                int typeValue = Optional.Check(type, 0);
                
                if (typeValue == 1)
                {
                    string base64Guid = Convert.ToBase64String(Guid.NewGuid().ToByteArray());

                    // Replace URL unfriendly characters
                    base64Guid = base64Guid.Replace('+', '-').Replace('/', '_');

                    // Remove the trailing ==
                    return base64Guid.Substring(0, base64Guid.Length - 2);
                }
                else
                {
                    return Guid.NewGuid().ToString();
                }

            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "String", Description = "Capitalize the first letter of the sentence.", HelpTopic = "Capitalize the first letter of the sentence.")]
        public static object SENTENCECASE(
[ExcelArgument("name", Name = "name", Description = "The name to convert.")] string name)
        {
            try
            {
                // start by converting entire string to lower case
                string lowerCase = name.ToLower();
                // matches the first sentence of a string, as well as subsequent sentences
                Regex r = new Regex(@"(^[a-z])|\.\s+(.)", RegexOptions.ExplicitCapture);
                // MatchEvaluator delegate defines replacement of setence starts to uppercase
                string result = r.Replace(lowerCase, s => s.Value.ToUpper());

                // result is: "This is a group. Of uncapitalized. Letters."
                return result;
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "String", Description = "Replaces the format items in a string with the string representation of three specified objects.", HelpTopic = "Replaces the format items in a string with the string representation of three specified objects.")]
        public static object FORMAT(
[ExcelArgument("format", Name = "format", Description = "A composite format string.")] string format,
[ExcelArgument("arg0", Name = "arg0", Description = "The first object to format.")] object arg0,
[ExcelArgument("[arg1]", Name = "[arg1]", Description = "(Optional) The second object to format.")] object arg1,
[ExcelArgument("[arg2]", Name = "[arg2]", Description = "(Optional) The third object to format.")] object arg2,
[ExcelArgument("[arg3]", Name = "[arg3]", Description = "(Optional) The fourth object to format.")] object arg3)
        {
            try
            {
                List<string> list0 = GetList(arg0);
                List<string> list1 = GetList(arg1);
                List<string> list2 = GetList(arg2);
                List<string> list3 = GetList(arg3);

                object[,] outputTable = Format(format, list0, list1,list2, list3);

                return outputTable;

            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        private static List<string> GetList(object arg)
        {
            List<string> list = new List<string>();

            if (arg is ExcelMissing) return list;
            if (arg is object[,])
            {
                object[,] argArray = (object[,])arg;

                foreach (string value in argArray)
                {
                    list.Add((string)value);
                }
            }
            else
            {
                list.Add(arg.ToString());
            }

            return list;
        }

        private static object[,] Format(string format, List<string> strings0, List<string> strings1, List<string> strings2, List<string> strings3)
        {
            List<string> formatedTexts = new List<string>();

            foreach (string value0 in strings0)
            {
                if (strings1.Count == 0)
                {
                    formatedTexts.Add(String.Format(format, value0));
                }
                foreach (string value1 in strings1)
                {
                    if (strings2.Count == 0)
                    {
                        formatedTexts.Add(String.Format(format, value0,value1));
                    }
                    foreach (string value2 in strings2)
                    {
                        if (strings3.Count == 0)
                        {
                            formatedTexts.Add(String.Format(format, value0, value1, value2));
                        }
                        foreach (string value3 in strings3)
                        {
                            formatedTexts.Add(String.Format(format, value0, value1, value2,value3));
                        }
                    }
                }
            }

            int l = 0;
            object[,] outputTable = new object[formatedTexts.Count, 1];

            foreach (string formatedText in formatedTexts)
            {
                outputTable[l,0] = formatedText;
                l++;
            }

            return outputTable;
        }
    }
}
