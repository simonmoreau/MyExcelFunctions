using ExcelDna.Integration;
using ExcelFunctions.Services;
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



        [ExcelFunction(Category = "String", Description = "Searches the specified input string for the specified regular expression", HelpTopic = "Searches the specified input string for the specified regular expression")]
        public static object REGEXEXTRACT(
            [ExcelArgument("input", Name = "input", Description = "The string to search for a match.")] string input,
            [ExcelArgument("pattern", Name = "pattern", Description = "The regular expression pattern to match.")] string pattern,
            [ExcelArgument("instance_num", Name = "instance_num", Description = "(Optional) A serial number that indicates which instance to extract. If omitted, returns the first found matches (default).")] object instance_num,
            [ExcelArgument("match_case", Name = "match_case", Description = "(Optional) Defines whether to match (TRUE or omitted) or ignore (FALSE) text case.")] object match_case)
        {
            try
            {
                bool _match_case = true;
                if (match_case != null)
                {
                    if (match_case.GetType() == typeof(ExcelDna.Integration.ExcelMissing))
                    {
                        _match_case = true;
                    }
                    else if (match_case.GetType() != typeof(bool))
                    {
                        return ExcelDna.Integration.ExcelError.ExcelErrorValue;
                    }
                    else
                    {
                        _match_case = Convert.ToBoolean(match_case);
                    }
                }


                int _instance_num = 0;
                if (instance_num != null)
                {
                    if (instance_num.GetType() == typeof(ExcelDna.Integration.ExcelMissing))
                    {
                        _instance_num = 0;
                    }
                    else if (instance_num.GetType() != typeof(double))
                    {
                        return ExcelDna.Integration.ExcelError.ExcelErrorValue;
                    }
                    else
                    {
                        _instance_num = Convert.ToInt32(instance_num);
                    }
                }


                Regex regex = new Regex(pattern);

                if (!_match_case)
                {
                    regex = new Regex(pattern, RegexOptions.IgnoreCase);
                }

                MatchCollection matches = regex.Matches(input);

                if (matches.Count > 0 && matches.Count > _instance_num)
                {
                    return matches[_instance_num].Value;
                }
                else
                {
                    return ExcelDna.Integration.ExcelError.ExcelErrorNull;
                }
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "String", Description = "Replaces the values matching a regex with the text you specify.", HelpTopic = "Replaces the values matching a regex with the text you specify.")]
        public static object REGEXREPLACE(
    [ExcelArgument("input", Name = "input", Description = "The string to search for a match.")] string input,
    [ExcelArgument("pattern", Name = "pattern", Description = "The regular expression pattern to match.")] string pattern,
    [ExcelArgument("replacement", Name = "replacement", Description = "The text to replace the matching substrings with.")] string replacement,
    [ExcelArgument("match_case", Name = "match_case", Description = "(Optional) Defines whether to match (TRUE or omitted) or ignore (FALSE) text case.")] object match_case)
        {
            try
            {
                bool _match_case = true;
                if (match_case != null)
                {
                    if (match_case.GetType() == typeof(ExcelDna.Integration.ExcelMissing))
                    {
                        _match_case = true;
                    }
                    else if (match_case.GetType() != typeof(bool))
                    {
                        return ExcelDna.Integration.ExcelError.ExcelErrorValue;
                    }
                    else
                    {
                        _match_case = Convert.ToBoolean(match_case);
                    }
                }

                Regex regex = new Regex(pattern);

                if (!_match_case)
                {
                    regex = new Regex(pattern, RegexOptions.IgnoreCase);
                }

                return regex.Replace(input, replacement);

            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }



        [ExcelFunction(Category = "String", Description = "Converts the string representation of a date and time to its System.DateTime equivalent.", HelpTopic = "Converts the string representation of a date and time to its System.DateTime equivalent.")]
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


    }
}
