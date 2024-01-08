using ExcelDna.Integration;
using ExcelFunctions.Services;
using HtmlAgilityPack;
using System.Collections.Specialized;
using ExcelDna.Registration.Utils;
using System.Net;
using System;

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

        [ExcelFunction(Category = "HttpUtility", Description = "Selects a list of nodes matching XPath expression from an HTML string.")]
        public static object SELECTNODES(
[ExcelArgument("html", Name = "html", Description = "String containing the HTML document to load. May not be null.")] string html,
[ExcelArgument("xpath", Name = "xpath", Description = "The XPath expression.")] string xpath,
[ExcelArgument("[attribute]", Name = "[attribute]", Description = "To get the given attribute value from the node.")] object attribute)
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

        [ExcelFunction(Category = "HttpUtility", Description = "Downloads the resource with the specified URI to a local file.")]
        public static object DOWNLOADFILE(
[ExcelArgument("uri", Name = "uri", Description = "The URI specified as a String, from which to download data.")] string uri,
[ExcelArgument("fileName", Name = "fileName", Description = "The name of the local file that is to receive the data.")] string fileName)
        {
            try
            {
                var functionName = nameof(DOWNLOADFILE);
                var parameters = new object[] { uri, fileName };

                return AsyncTaskUtil.RunTask<object>(functionName, parameters, async () =>
                {
                    using HttpClient client = new HttpClient();
                    using Stream stream = await client.GetStreamAsync(uri);
                    using FileStream fileStream = new FileStream(fileName, FileMode.OpenOrCreate);
                    await stream.CopyToAsync(fileStream);

                    return fileName;
                });
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }
    }
}

