using ExcelDna.Integration;
using ExcelDna.Serialization;
using ExcelFunctions.Services;
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
[ExcelArgument("table", Name = "table", Description = "The table of values to convert.")] object values)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;

                    Type type = null;
                    List<object> objects = ObjectsListBuilder.BuilObjectList(inputArray, out type);

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

        [ExcelFunction(Category = "Serialization", Description = "Converts the provided value into a XML string.", HelpTopic = "Converts the provided value into a XML string.")]
        public static object XMLSERIALIZE(
[ExcelArgument("table", Name = "table", Description = "The table of values to convert.")] object values)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;

                    Type objectType = null;
                    List<object> objects = ObjectsListBuilder.BuilObjectList(inputArray, out objectType);

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
    }
}
