using ExcelDna.Integration;
using ExcelFunctions.Services;
using HtmlAgilityPack;
using System.Collections.Specialized;
using ExcelDna.Registration.Utils;
using System.Net;
using System;
using System.IO;

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


                string functionName = nameof(SELECTHTMLNODES);
                object[] parameters = new object[] { url, xpath };

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
                HtmlDocument doc = new HtmlDocument();
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
[ExcelArgument("uri", Name = "uri", Description = "The URI specified as a String, from which to download data.")] object uri,
[ExcelArgument("directory", Name = "directory", Description = "The directory to save the data.")] string directory,
[ExcelArgument("[fileName]", Name = "[fileName]", Description = "The name of the local file that is to receive the data.")] object fileName)
        {
            try
            {
                string functionName = nameof(DOWNLOADFILE);
                object[] parameters = new object[] { uri, directory, fileName };

                return AsyncTaskUtil.RunTask<object>(functionName, parameters, async () =>
                {
                    List<string> uris = new List<string>();
                    if (uri.GetType() == typeof(string))
                    {
                        uris.Add((string)uri);
                    }
                    else if (uri.GetType() == typeof(object[,]))
                    {
                        object[,] inputArray = (object[,])uri;
                        string?[] inputColumn = Enumerable.Range(0, inputArray.GetLength(0)).Select(x => Convert.ToString(inputArray[x, 0])).ToArray();
                        uris.AddRange(inputColumn);
                    }
                    else
                    {
                        return ExcelDna.Integration.ExcelError.ExcelErrorNA;
                    }

                    string directoryPath = Directory.CreateDirectory(directory).FullName;

                    List<Task> tasks = new List<Task>();

                    foreach (string url in uris)
                    {
                        string urlFileName = Path.GetFileName(url);
                        string filePath = Path.Combine(directoryPath, urlFileName);
                        string fileNameValue = Optional.Check(fileName, "");

                        if (fileNameValue != "")
                        {
                            filePath = Path.Combine(directoryPath, fileNameValue);
                        }

                        tasks.Add(DownloadFile(url, filePath));
                    }

                    await Task.WhenAll(tasks);

                    return directoryPath;
                });
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        private static async Task DownloadFile(string uri, string filePath)
        {
            try
            {
                using HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Add("User-Agent", "ExcelDownloadBot/0.0 (https://www.bim42.com/; simon@bim42.com)");
                using Stream stream = await client.GetStreamAsync((string)uri);
                using FileStream fileStream = new FileStream(filePath, FileMode.OpenOrCreate);
                await stream.CopyToAsync(fileStream);
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        [ExcelFunction(Category = "HttpUtility", Description = "Encodes a URL string.")]
        public static object URLENCODE(
[ExcelArgument("str", Name = "str", Description = "The text to encode.")] string str)
        {
            try
            {
                return System.Web.HttpUtility.UrlEncode(str);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "HttpUtility", Description = "Converts a string that has been encoded for transmission in a URL into a decoded string.")]
        public static object URLDECODE(
[ExcelArgument("str", Name = "str", Description = "The string to decode.")] string str)
        {
            try
            {
                return System.Web.HttpUtility.UrlDecode(str);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }
    }

}

