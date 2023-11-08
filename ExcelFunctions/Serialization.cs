using ExcelDna.Integration;
using ExcelFunctions.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

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

                    List<object> objects = ObjectsListBuilder.BuilObjectList(inputArray);

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
    }
}
