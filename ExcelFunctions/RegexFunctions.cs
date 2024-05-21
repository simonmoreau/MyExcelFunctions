using ExcelDna.Integration;
using ExcelFunctions.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelFunctions
{
    public static class RegexFunctions
    {
        [ExcelFunction(Category = "Regex", Description = "Searches the specified input string for all occurrences of a regular expression.", HelpTopic = "Searches the specified input string for all occurrences of a regular expression.")]
        public static object REGEXMATCHES(
    [ExcelArgument("input", Name = "input", Description = "The string to search for a match.")] string input,
    [ExcelArgument("pattern", Name = "pattern", Description = "The regular expression pattern to match.")] string pattern,
    [ExcelArgument("[match_case]", Name = "[match_case]", Description = "(Optional) Defines whether to match (TRUE or omitted) or ignore (FALSE) text case.")] object match_case,
    [ExcelArgument("[capturing_group]", Name = "[capturing_group]", Description = "(Optional) Return the results from a single capturing group. Can either be an integer or a string in the case of a named capturing group")] object capturing_group)
        {
            try
            {

                bool _match_case = Optional.Check(match_case, true);

                string? namedCapturingGroup = null;
                if (capturing_group is string)
                {
                    namedCapturingGroup = capturing_group as string;
                }

                int? capturingGroupIndex = null;
                if (capturing_group is double)
                {
                    capturingGroupIndex = Convert.ToInt32(capturing_group);
                }

                Regex regex = new Regex(pattern);

                if (!_match_case)
                {
                    regex = new Regex(pattern, RegexOptions.IgnoreCase);
                }

                MatchCollection matches = regex.Matches(input);

                if (matches.Count == 0)
                {
                    return ExcelDna.Integration.ExcelError.ExcelErrorNull;
                }

                // return matches[_instance_num].Value;

                object[,] outputTable = new object[matches.Count, 1];

                int l = 0;
                foreach (Match match in matches)
                {
                    if (namedCapturingGroup != null)
                    {
                        outputTable[l, 0] = match.Groups[namedCapturingGroup].Value;
                    }
                    else if (capturingGroupIndex != null)
                    {
                        outputTable[l, 0] = match.Groups[capturingGroupIndex.Value].Value;
                    }
                    else
                    {
                        outputTable[l, 0] = match.Value;
                    }
                    l++;
                }
                return outputTable;

            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
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

        [ExcelFunction(Category = "Regex", Description = "Replaces the values matching a regex with the text you specify.", HelpTopic = "Replaces the values matching a regex with the text you specify.")]
        public static object REGEXREPLACE(
    [ExcelArgument("input", Name = "input", Description = "The string to search for a match.")] string input,
    [ExcelArgument("pattern", Name = "pattern", Description = "The regular expression pattern to match.")] string pattern,
    [ExcelArgument("replacement", Name = "replacement", Description = "The text to replace the matching substrings with.")] string replacement,
    [ExcelArgument("match_case", Name = "match_case", Description = "(Optional) Defines whether to match (TRUE or omitted) or ignore (FALSE) text case.")] object match_case)
        {
            try
            {
                bool _match_case = Optional.Check(match_case, true);

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

    }
}
