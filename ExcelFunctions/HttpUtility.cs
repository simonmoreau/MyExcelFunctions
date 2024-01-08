using ExcelDna.Integration;
using ExcelFunctions.Services;
using HtmlAgilityPack;
using System.Collections.Specialized;
using ExcelDna.Registration.Utils;

namespace ExcelFunctions
{
    public static class HttpUtility
    {
        [ExcelFunction(Category = "HttpUtility", Description = "Changes the extension of a path string.")]
        public static object PARSEQUERYSTRING(
    [ExcelArgument("query", Name = "query", Description = "The query string to parse.")] string query,
    [ExcelArgument("parameterName", Name = "parameterName", Description = "The name of the query string parameter to retrieve.")] string parameterName)
        {
            try
            {
                NameValueCollection nameValueCollection = System.Web.HttpUtility.ParseQueryString(query);
                if (nameValueCollection.AllKeys.Contains(parameterName))
                {
                    return nameValueCollection[parameterName];
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

        [ExcelFunction(Category = "HttpUtility", Description = "Selects a list of nodes matching XPath expression from an Internet resource.")]
        public static object SELECTHTMLNODES(
[ExcelArgument("url", Name = "url", Description = "The requested URL, such as \"http://Myserver/Mypath/Myfile.asp\".")] string url,
[ExcelArgument("xpath", Name = "xpath", Description = "The XPath expression.")] string xpath)
        {
            try
            {
                // From Web
                HtmlWeb web = new HtmlWeb();
                

                var functionName = nameof(SELECTHTMLNODES);
                var parameters = new object[] { url, xpath };

                return AsyncTaskUtil.RunTask<object>(functionName, parameters, async () =>
                {
                    //The actual asyncronous block of code to execute.

                    HtmlDocument doc = await web.LoadFromWebAsync(url);

                    HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes(xpath);

                    if (nodes.Count == 0) return ExcelDna.Integration.ExcelError.ExcelErrorNull;

                    object[,] outputTable = new object[nodes.Count, 1];

                    int l = 0;
                    foreach (HtmlNode? node in nodes)
                    {
                        outputTable[l, 0] = node.OuterHtml;
                        l++;
                    }
                    return outputTable;
                });
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "HttpUtility", Description = "Selects the first node matching XPath expression from an HTML string.")]
        public static object SELECTNODE(
[ExcelArgument("html", Name = "html", Description = "String containing the HTML document to load. May not be null.")] string html,
[ExcelArgument("xpath", Name = "xpath", Description = "The XPath expression.")] string xpath,
[ExcelArgument("[attribute]", Name = "[attribute]", Description = "The minimun score to get a match.")] object attribute)
        {
            try
            {
                // From String
                var doc = new HtmlDocument();
                doc.LoadHtml(html);

                string attribueName = Optional.Check(attribute, "");

                HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes(xpath);

                if (nodes.Count == 0) return ExcelDna.Integration.ExcelError.ExcelErrorNull;

                object[,] outputTable = new object[nodes.Count, 1];

                int l = 0;
                foreach (HtmlNode? node in nodes)
                {
                    if (attribueName == "")
                    {
                        outputTable[l, 0] = node.InnerHtml;
                    }
                    else
                    {
                        outputTable[l, 0] = node.Attributes[attribueName].Value;
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
    }
}

