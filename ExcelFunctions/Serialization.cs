using ExcelDna.Integration;
using ExcelDna.Serialization;
using ExcelFunctions.Services;
using ExcelFunctions.XML;
using Microsoft.Extensions.FileProviders;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Json.Path;
using System.Text.Json.Nodes;

namespace ExcelFunctions
{
    public static class Serialization
    {
        [ExcelFunction(Category = "Serialization", Description = "Converts the provided value into a JSON string.", HelpTopic = "Converts the provided value into a JSON string.")]
        public static object JSONSERIALIZE(
[ExcelArgument("table", Name = "table", Description = "The table of values to convert.")] object values,
[ExcelArgument("[dates]", Name = "[dates]", Description = "[optional]A comma-separated list of column name to be serialized as dates")] object dates)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;

                    Type type = null;
                    string? datesColumnsName = Optional.Check(dates,"");
                    string[]? datesColumns = null;
                    if (!string.IsNullOrEmpty(datesColumnsName))
                    {
                        datesColumns = datesColumnsName.Split(',').Select(s => s.Trim()).ToArray();
                    }
                    List<object> objects = ObjectsListBuilder.BuilObjectList(inputArray, out type, datesColumns);

                    JsonSerializerOptions options = new JsonSerializerOptions
                    {
                        WriteIndented = true
                    };

                    string jsonString = JsonSerializer.Serialize(objects, options);
                    return jsonString;

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

        // NEW: Get a property from a JSON string using a JsonPath expression
        [ExcelFunction(Category = "Serialization", Description = "Gets value(s) from a JSON string by a JsonPath expression.", HelpTopic = "Extracts JSON value(s) using JsonPath.")]
        public static object JSONGETPATH(
            [ExcelArgument("json", Name = "json", Description = "The JSON text to query.")] string jsonText,
            [ExcelArgument("path", Name = "path", Description = "JsonPath expression (e.g. $.store.book[0].title)")] string jsonPathExpr)
        {
            try
            {
                if (string.IsNullOrEmpty(jsonText) || string.IsNullOrEmpty(jsonPathExpr))
                    return ExcelDna.Integration.ExcelError.ExcelErrorNA;

                JsonPath path = JsonPath.Parse(jsonPathExpr);
                JsonNode? document = JsonNode.Parse(jsonText);

                PathResult matches = path.Evaluate(document);


                List<string> results = new List<string>();
                if (matches.Matches is System.Collections.IEnumerable enumerable)
                {
                    foreach (Json.Path.Node? m in enumerable)
                    {
                        if (m == null) continue;
                        results.Add(m.Value?.ToString());
                    }
                }
                else
                {
                    results.Add(matches?.ToString());
                }

                if (results.Count == 0 || (results.Count == 1 && string.IsNullOrEmpty(results[0])))
                    return ExcelDna.Integration.ExcelError.ExcelErrorNA;

                if (results.Count == 1)
                    return results[0];

                // Return as a vertical 2D array: one column, multiple rows
                object[,] output = new object[results.Count, 1];
                for (int i = 0; i < results.Count; i++)
                {
                    output[i, 0] = results[i];
                }
                return output;
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }
        [ExcelFunction(Category = "Serialization", Description = "Converts the provided value into a JSON file.", HelpTopic = "Converts the provided value into a JSON file.")]
        public static object JSONSERIALIZETOFILE(
                [ExcelArgument("table", Name = "table", Description = "The table of values to convert.")] object values,
                [ExcelArgument("path", Name = "path", Description = "The path to the json file")] object path,
                [ExcelArgument("[dates]", Name = "[dates]", Description = "[optional]A comma-separated list of column name to be serialized as dates")] object dates)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;

                    Type type = null;
                    string? datesColumnsName = Optional.Check(dates, "");
                    string[]? datesColumns = null;
                    if (!string.IsNullOrEmpty(datesColumnsName))
                    {
                        datesColumns = datesColumnsName.Split(',').Select(s => s.Trim()).ToArray();
                    }
                    List<object> objects = ObjectsListBuilder.BuilObjectList(inputArray, out type, datesColumns);

                    JsonSerializerOptions options = new JsonSerializerOptions
                    {
                        WriteIndented = true
                    };

                    string jsonString = JsonSerializer.Serialize(objects, options);

                    if (path == null) throw new NullReferenceException("The path is null");

                    string stringPath = path.ToString();

                    File.WriteAllText(stringPath, jsonString);

                    return $"The file has be written to {stringPath}";

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

        [ExcelFunction(Category = "Serialization", Description = "Converts the provided value into a XML string.", HelpTopic = "Converts the provided value into a XML string.")]
        public static object XMLSERIALIZE(
[ExcelArgument("table", Name = "table", Description = "The table of values to convert.")] object values,
[ExcelArgument("[dates]", Name = "[dates]", Description = "[optional]A list of column to be serialized as dates")] object dates)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;

                    Type objectType = null;
                    string[]? datesColumns = Optional.Check(dates);
                    List<object> objects = ObjectsListBuilder.BuilObjectList(inputArray, out objectType, datesColumns);

                    Type listType = typeof(List<>).MakeGenericType(objectType);
                    IList typedList = (IList)Activator.CreateInstance(listType);

                    foreach (object item in objects)
                    {
                        typedList.Add(item);
                    }

                    XmlSerializer xmlSerializer = new XmlSerializer(listType);

                    using (StringWriter textWriter = new StringWriter())
                    {
                        xmlSerializer.Serialize(textWriter, typedList);
                        return textWriter.ToString();
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

        [ExcelFunction(Category = "Serialization", Description = "Converts the provided value into a XML file.", HelpTopic = "Converts the provided value into a XML string.")]
        public static object XMLSERIALIZETOFILE2(
                [ExcelArgument("table", Name = "table", Description = "The table of values to convert.")] object values,
                [ExcelArgument("path", Name = "path", Description = "The path to the XML file")] object path,
                [ExcelArgument("[dates]", Name = "[dates]", Description = "[optional]A list of column to be serialized as dates")] object dates)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;

                    Type objectType = null;
                    string[]? datesColumns = Optional.Check(dates);
                    List<object> objects = ObjectsListBuilder.BuilObjectList(inputArray, out objectType, datesColumns);

                    Type listType = typeof(List<>).MakeGenericType(objectType);
                    IList typedList = (IList)Activator.CreateInstance(listType);

                    foreach (object item in objects)
                    {
                        typedList.Add(item);
                    }

                    XmlSerializer xmlSerializer = new XmlSerializer(listType);

                    using (StringWriter textWriter = new StringWriter())
                    {
                        xmlSerializer.Serialize(textWriter, typedList);

                        if (path == null) throw new NullReferenceException("The path is null");

                        string stringPath = path.ToString();

                        File.WriteAllText(stringPath, textWriter.ToString());

                        return $"The file has be written to {stringPath}";
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


    }
}
