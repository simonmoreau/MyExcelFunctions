using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace MyExcelFunctions
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
    }
}

