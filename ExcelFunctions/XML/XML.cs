using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace ExcelFunctions.XML
{
    public static class XML
    {
        [ExcelFunction(Category = "My functions", Description = "Convert a range to an XML string.", HelpTopic = "Convert a range to an XML string. The first line are the fields of the XML file")]
        public static object XMLSERIALIZE(
[ExcelArgument("table", Name = "table", Description = "The table to convert to XML")] object values,
[ExcelArgument("Name", Name = "Name", Description = "The name of each row")] object rowName)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;

                    string xml = GetXmlString(ref rowName, inputArray);

                    return xml;

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

        [ExcelFunction(Category = "My functions", Description = "Convert a range to an XML string.", HelpTopic = "Convert a range to an XML string. The first line are the fields of the XML file")]
        public static object XMLSERIALIZETOFILE(
[ExcelArgument("table", Name = "table", Description = "The table to convert to XML")] object values,
[ExcelArgument("path", Name = "path", Description = "The path to the xml file")] object path,
[ExcelArgument("Name", Name = "Name", Description = "The name of each row")] object rowName)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;

                    string xml = GetXmlString(ref rowName, inputArray);

                    if (path != null)
                    {
                        string stringPath = path.ToString();
                        if (System.IO.File.Exists(stringPath))
                        {
                            File.WriteAllText(stringPath, xml.Replace("ObjectSerialize", "NewDataSet"));
                        }
                    }

                    return "The file has be written";

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

        private static string GetXmlString(ref object rowName, object[,] inputArray)
        {
            // Build an dictonary of field

            List<DynamicField> fields = new List<DynamicField>();
            for (int i = 0; i < inputArray.GetLength(1); i++)
            {
                string name = inputArray[0, i].ToString();
                Type fieldType = GetNullableType(inputArray[1, i].GetType());

                int rowIndex = 1;
                while (fieldType.FullName == "ExcelDna.Integration.ExcelEmpty" && rowIndex < inputArray.GetLength(0))
                {
                    fieldType = GetNullableType(inputArray[rowIndex, i].GetType());
                    rowIndex++;
                }
                fields.Add(new DynamicField(name, fieldType));
            }

            if (rowName.GetType() == typeof(ExcelDna.Integration.ExcelMissing))
            {
                rowName = "object";
            }

            // Create a new type to be exported
            Type type = XmlTypeBuilder.CompileResultType(fields.ToArray(), rowName.ToString());

            // Create a list of object of this type
            List<object> objects = new List<object>();

            for (int i = 1; i < inputArray.GetLength(0); i++)
            {
                object rowObject = Activator.CreateInstance(type);

                for (int j = 0; j < inputArray.GetLength(1); j++)
                {
                    PropertyInfo propertyInfo = rowObject.GetType().GetProperty(fields[j].Name);
                    if (GetNullableType(inputArray[i, j].GetType()) == fields[j].Type)
                    {
                        propertyInfo.SetValue(rowObject, inputArray[i, j], null);
                    }
                    else if (inputArray[i, j].GetType().FullName == "ExcelDna.Integration.ExcelEmpty")
                    {
                        propertyInfo.SetValue(rowObject, null, null);
                    }
                    else
                    {
                        object castedObject = null;
                        try
                        {
                            castedObject = Convert.ChangeType(inputArray[i, j], fields[j].Type);
                        }
                        catch
                        {
                        }
                        propertyInfo.SetValue(rowObject, castedObject, null);
                    }
                }

                objects.Add(rowObject);
            }

            var objectSerialize = new ObjectSerialize
            {
                ObjectList = objects
            };

            XmlSerializer xsSubmit = new XmlSerializer(typeof(ObjectSerialize));
            var xml = "";

            using (var sww = new StringWriter())
            {
                using (XmlWriter writers = XmlWriter.Create(sww))
                {
                    xsSubmit.Serialize(writers, objectSerialize);
                    xml = sww.ToString(); // Your XML
                    xml = FormatXml(xml);
                }
            }

            return xml;
        }

        static string FormatXml(string xml)
        {
            try
            {
                XDocument doc = XDocument.Parse(xml);
                return doc.ToString();
            }
            catch (Exception)
            {
                // Handle and throw if fatal exception here; don't just ignore them
                return xml;
            }
        }

        private static Type GetNullableType(Type type)
        {
            // Use Nullable.GetUnderlyingType() to remove the Nullable<T> wrapper if type is already nullable.
            type = Nullable.GetUnderlyingType(type) ?? type; // avoid type becoming null
            if (type.IsValueType)
                return typeof(Nullable<>).MakeGenericType(type);
            else
                return type;
        }

    }
}
