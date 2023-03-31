using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MyExcelFunctions
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

    }
}
