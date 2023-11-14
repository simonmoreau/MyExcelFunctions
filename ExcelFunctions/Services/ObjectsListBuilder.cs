using ExcelFunctions.XML;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Xml;
using System.Reflection.Emit;
using System.Data;
using System.Collections;

namespace ExcelFunctions.Services
{
    internal class ObjectsListBuilder
    {

        public static List<object> BuilObjectList(object[,] inputArray)
        {
            Dictionary<string, Type> columnsWithType = ListColumns(inputArray);

            Dictionary<string, object> root = new Dictionary<string, object>();

            foreach (KeyValuePair<string, Type> dotObject in columnsWithType)
            {
                string[] hierarcy = dotObject.Key.Split('.');

                Dictionary<string, object> current = root;

                for (int i = 0; i < hierarcy.Length; i++)
                {
                    string key = hierarcy[i];

                    if (i == hierarcy.Length - 1) // Last key
                    {
                        current.Add(key, dotObject.Value);
                    }
                    else
                    {
                        if (!current.ContainsKey(key))
                        {
                            current.Add(key, new Dictionary<string, object>());
                        }

                        current = (Dictionary<string, object>)current[key];
                    }
                }
            }

            Type buildedType = BuildType(root, "root");

            Dictionary<int, int> columnsRanks = new Dictionary<int, int>();

            int index = 0;
            foreach (string columnName in columnsWithType.Keys)
            {
                int rank = 0;
                if (columnName.Contains("."))
                {
                    rank = columnName.Count(f => f == '.');
                }
                columnsRanks.Add(index, rank);
                index++;
            }

            List<List<object>> inputLists = new List<List<object>>();

            for (int i = 1; i < inputArray.GetLength(0); i++)
            {
                List<object> row = new List<object>();
                for (int j = 0; j < inputArray.GetLength(1); j++)
                {
                    row.Add(inputArray[i, j]);
                }

                inputLists.Add(row);
            }

            int maxRank = columnsRanks.Values.Max();
            for (int i = 0; i < maxRank; i++)
            {

            }
            GroupInputRow(columnsRanks, inputLists);




            Dictionary<string, object> rowObjects = new Dictionary<string, object>();

            foreach (List<object> inputList in inputLists)
            {
                object rowObject = null;
                string groupingKey = GroupingString(inputList, columnsRanks, 0);
                if (rowObjects.ContainsKey(groupingKey))
                {
                    rowObject = rowObjects[groupingKey];
                }
                else
                {
                    rowObject = Activator.CreateInstance(buildedType);
                    rowObjects.Add(groupingKey, rowObject);
                }


                for (int j = 0; j < inputList.Count; j++)
                {
                    string propertyPath = columnsWithType.ElementAt(j).Key;

                    if (GetNullableType(inputList[j].GetType()) == columnsWithType.ElementAt(j).Value)
                    {
                        SetPropertyValue(rowObject, propertyPath, inputList[j]);
                    }
                    else if (inputList[j].GetType().FullName == "ExcelDna.Integration.ExcelEmpty")
                    {
                        SetPropertyValue(rowObject, propertyPath, null);
                    }
                    else
                    {
                        object castedObject = null;
                        try
                        {
                            castedObject = Convert.ChangeType(inputList[j], columnsWithType.ElementAt(j).Value);
                        }
                        catch
                        {
                        }
                        SetPropertyValue(rowObject, propertyPath, castedObject);
                    }
                }

            }

            // Create a list of object of this type
            List<object> objects = new List<object>();

            objects.AddRange(rowObjects.Values);

            //for (int i = 1; i < inputArray.GetLength(0); i++)
            //{
            //    object rowObject = Activator.CreateInstance(buildedType);

            //    for (int j = 0; j < inputArray.GetLength(1); j++)
            //    {
            //        string propertyPath = columnsWithType.ElementAt(j).Key;

            //        if (GetNullableType(inputArray[i, j].GetType()) == columnsWithType.ElementAt(j).Value)
            //        {
            //            SetPropertyValue(rowObject, propertyPath, inputArray[i, j]);
            //        }
            //        else if (inputArray[i, j].GetType().FullName == "ExcelDna.Integration.ExcelEmpty")
            //        {
            //            SetPropertyValue(rowObject, propertyPath, null);
            //        }
            //        else
            //        {
            //            object castedObject = null;
            //            try
            //            {
            //                castedObject = Convert.ChangeType(inputArray[i, j], columnsWithType.ElementAt(j).Value);
            //            }
            //            catch
            //            {
            //            }
            //            SetPropertyValue(rowObject, propertyPath, castedObject);
            //        }
            //    }

            //    objects.Add(rowObject);
            //}



            return objects;
        }

        //private static Dictionary<List<ObjectGrouping>> GroupRows(Dictionary<string,List<object>> groupedRow, Dictionary<int, int> columnsRanks)
        //{
        //    IEnumerable<IGrouping<string, List<object>>> groups = groupedRow.Values.GroupBy(r => GroupingString(r, columnsRanks, 0));

        //    foreach (IGrouping<string, List<object>> group in groups)
        //    {
        //        IEnumerable<IGrouping<string, List<object>>> test = group.GroupBy(r => GroupingString(r, columnsRanks, 1));
        //    }
        //}
        private static void GroupInputRow(Dictionary<int, int> columnsRanks, List<List<object>> inputList)
        {
            IEnumerable<IGrouping<string, List<object>>> groups = inputList.GroupBy(r => GroupingString(r, columnsRanks, 0));

            Dictionary<string, Dictionary<string, List<object>>> groupedRow = new Dictionary<string, Dictionary<string, List<object>>>();

            foreach (IGrouping<string, List<object>> group in groups)
            {
                //ObjectGrouping objectGrouping = new ObjectGrouping();
                //objectGrouping.Name = group.Key;
                //objectGrouping.ObjectGroupings.Add(objectGrouping);
                //groupedRow.Add(group.Key, GroupRows()
                //IEnumerable < IGrouping<string, List<object>> > test = group.GroupBy(r => GroupingString(r, columnsRanks, 1));
            }
        }

        private static string GroupingString(List<object> row, Dictionary<int, int> columnsRanks, int rank)
        {
            string groupingString = "";
            foreach (KeyValuePair<int, int> indexRank in columnsRanks)
            {
                if (indexRank.Value == rank)
                {
                    string value = row[indexRank.Key].ToString();
                    if (row[indexRank.Key].GetType().FullName == "ExcelDna.Integration.ExcelEmpty")
                    {
                        value = "";
                    }
                    groupingString = groupingString + ";" + value;
                }
            }

            return groupingString;
        }

        private static Dictionary<string, Type> ListColumns(object[,] inputArray)
        {
            // Build an dictonary of field
            Dictionary<string, Type> columnsWithType = new Dictionary<string, Type>();


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

                columnsWithType.Add(name, fieldType);
            }

            return columnsWithType;
        }

        private static object GetPropertyValue(object obj, string propertyName)
        {
            foreach (PropertyInfo propertyInfo in propertyName.Split('.').Select(s => obj.GetType().GetProperty(s)))
            {
                obj = propertyInfo.GetValue(obj, null);
            }
            return obj;
        }

        private static void SetPropertyValue(object parentTarget, string compoundProperty, object value)
        {
            string[] bits = compoundProperty.Split('.');
            for (int i = 0; i < bits.Length - 1; i++)
            {
                PropertyInfo propertyToGet = parentTarget.GetType().GetProperty(bits[i]);

                object target = propertyToGet.GetValue(parentTarget, null);

                if (target == null)
                {
                    target = Activator.CreateInstance(propertyToGet.PropertyType);
                    propertyToGet.SetValue(parentTarget, target);
                }

                parentTarget = target;
            }
            PropertyInfo propertyToSet = parentTarget.GetType().GetProperty(bits.Last());
            propertyToSet.SetValue(parentTarget, value, null);
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

        private static TypeBuilder GetTypeBuilder(string typeSignature)
        {
            AssemblyName an = new AssemblyName(typeSignature);
            AssemblyBuilder assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(new AssemblyName(Guid.NewGuid().ToString()), AssemblyBuilderAccess.Run);
            ModuleBuilder moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");
            TypeBuilder tb = moduleBuilder.DefineType(typeSignature,
                    TypeAttributes.Public |
                    TypeAttributes.Class |
                    TypeAttributes.AutoClass |
                    TypeAttributes.AnsiClass |
                    TypeAttributes.BeforeFieldInit |
                    TypeAttributes.AutoLayout,
                    null);
            return tb;
        }

        private static void CreateProperty(TypeBuilder tb, string propertyName, Type propertyType)
        {
            FieldBuilder fieldBuilder = tb.DefineField("_" + propertyName, propertyType, FieldAttributes.Private);

            PropertyBuilder propertyBuilder = tb.DefineProperty(propertyName, PropertyAttributes.HasDefault, propertyType, null);
            MethodBuilder getPropMthdBldr = tb.DefineMethod("get_" + propertyName, MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig, propertyType, Type.EmptyTypes);
            ILGenerator getIl = getPropMthdBldr.GetILGenerator();

            getIl.Emit(OpCodes.Ldarg_0);
            getIl.Emit(OpCodes.Ldfld, fieldBuilder);
            getIl.Emit(OpCodes.Ret);

            MethodBuilder setPropMthdBldr =
                tb.DefineMethod("set_" + propertyName,
                  MethodAttributes.Public |
                  MethodAttributes.SpecialName |
                  MethodAttributes.HideBySig,
                  null, new[] { propertyType });

            ILGenerator setIl = setPropMthdBldr.GetILGenerator();
            Label modifyProperty = setIl.DefineLabel();
            Label exitSet = setIl.DefineLabel();

            setIl.MarkLabel(modifyProperty);
            setIl.Emit(OpCodes.Ldarg_0);
            setIl.Emit(OpCodes.Ldarg_1);
            setIl.Emit(OpCodes.Stfld, fieldBuilder);

            setIl.Emit(OpCodes.Nop);
            setIl.MarkLabel(exitSet);
            setIl.Emit(OpCodes.Ret);

            propertyBuilder.SetGetMethod(getPropMthdBldr);
            propertyBuilder.SetSetMethod(setPropMthdBldr);
        }

        private static Type BuildType(Dictionary<string, object> fields, string name)
        {
            TypeBuilder tb = GetTypeBuilder(name);
            ConstructorBuilder constructor = tb.DefineDefaultConstructor(MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.RTSpecialName);

            foreach (KeyValuePair<string, object> field in fields)
            {
                if (field.Value.GetType().FullName == "System.RuntimeType")
                {
                    CreateProperty(tb, field.Key, (Type)field.Value);
                }
                else
                {
                    Dictionary<string, object> nestedType = (Dictionary<string, object>)field.Value;
                    // Type genericListType = typeof(List<>).MakeGenericType(BuildType(nestedType, field.Key));
                    CreateProperty(tb, field.Key, BuildType(nestedType, field.Key));
                }
            }

            Type objectType = tb.CreateType();
            return objectType;
        }
    }

    public class ObjectGrouping
    {
        public ObjectGrouping()
        {
            ObjectGroupings = new List<ObjectGrouping>();
        }
        public string Name { get; set; }
        public List<ObjectGrouping> ObjectGroupings { get; set; }
    }
}
