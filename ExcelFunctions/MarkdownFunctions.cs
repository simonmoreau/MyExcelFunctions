using ExcelDna.Integration;
using ExcelFunctions.Services;
using FuzzySharp;
using Markdig;
using Microsoft.Extensions.FileSystemGlobbing.Internal;
using System;
using System.Data.Common;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ExcelFunctions
{
    public static class MarkdownFunctions
    {

        [ExcelFunction(Category = "String", Description = "Converts a Markdown string to HTML.", HelpTopic = "Converts a Markdown string to HTML.")]
        public static object HTML(
    [ExcelArgument("markdown", Name = "markdown", Description = "A Markdown text.")] string markdown)
        {
            try
            {
                string result = Markdown.ToHtml(markdown);

                return result;
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }
    }
}
