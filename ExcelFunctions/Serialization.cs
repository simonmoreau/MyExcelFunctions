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
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ExcelFunctions
{
    public static class Serialization
    {
        [ExcelFunction(Category = "Serialization", Description = "Converts the provided value into a JSON string.", HelpTopic = "Converts the provided value into a JSON string.")]
        public static object JSONSERIALIZE(
[ExcelArgument("table", Name = "table", Description = "The table of values to convert.")] object values,
[ExcelArgument("[dates]", Name = "[dates]", Description = "[optional]A list of column to be serialized as dates")] object dates)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;

                    Type type = null;
                    string[]? datesColumns = Optional.Check(dates);
                    List<object> objects = ObjectsListBuilder.BuilObjectList(inputArray, out type, datesColumns);

                    string jsonString = JsonSerializer.Serialize(objects);
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

        [ExcelFunction(Category = "Serialization", Description = "Converts the provided value into a JSON file.", HelpTopic = "Converts the provided value into a JSON file.")]
        public static object JSONSERIALIZETOFILE(
                [ExcelArgument("table", Name = "table", Description = "The table of values to convert.")] object values,
                [ExcelArgument("path", Name = "path", Description = "The path to the json file")] object path,
                [ExcelArgument("[dates]", Name = "[dates]", Description = "[optional]A list of column to be serialized as dates")] object dates)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;

                    Type type = null;
                    string[]? datesColumns = Optional.Check(dates);
                    List<object> objects = ObjectsListBuilder.BuilObjectList(inputArray, out type, datesColumns);

                    string jsonString = JsonSerializer.Serialize(objects);

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
