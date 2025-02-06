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
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace ExcelFunctions.Services
{
    internal class PropertyHolder
    {
        public Dictionary<string, object> Fields = new Dictionary<string, object>();
        public Dictionary<string, List<PropertyHolder>> Properties = new Dictionary<string, List<PropertyHolder>>();
    }
    internal class ObjectsListBuilder
    {

        public static List<object> BuilObjectList(object[,] inputArray, out Type buildedType, string[]? dateColumns = null)
        {
            List<Dictionary<string, object>> dataTree = CreateDataTree(inputArray, dateColumns);

            IEnumerable<IGrouping<string, Dictionary<string, object>>> groups = dataTree.GroupBy(d => GroupingString(d));

            List<PropertyHolder> propertyHolders = new List<PropertyHolder>();

            foreach (IGrouping<string, Dictionary<string, object>> group in groups)
            {
                propertyHolders.Add(BuildHolder(group));
            }

            Dictionary<string, Type> columnsWithType = ListColumns(inputArray, dateColumns);

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

            buildedType = BuildType(root, "root");
            List<object> objects = new List<object>();

            foreach (PropertyHolder propertyHolder in propertyHolders)
            {
                object rowObject = Activator.CreateInstance(buildedType);

                int[] indexes = new int[columnsWithType.Count];
                ProcessProperty(propertyHolder, rowObject, "", indexes, 0);

                objects.Add(rowObject);
            }


            return objects;
        }

        private static void ProcessProperty(PropertyHolder propertyHolder, object rowObject, string basePropertyPath, int[] indexes, int rank)
        {
            foreach (KeyValuePair<string, object> field in propertyHolder.Fields)
            {
                string propertyPath = basePropertyPath + "." + field.Key;
                if (basePropertyPath == "")
                {
                    propertyPath = field.Key;
                }

                SetPropertyValue(rowObject, propertyPath, field.Value, indexes, rank);

            }

            foreach (KeyValuePair<string, List<PropertyHolder>> properties in propertyHolder.Properties)
            {
                int i = 0;
                foreach (PropertyHolder property in properties.Value)
                {
                    string propertyPath = basePropertyPath + "." + properties.Key;
                    if (basePropertyPath == "")
                    {
                        propertyPath = properties.Key;
                    }

                    indexes[rank + 1] = i;

                    ProcessProperty(property, rowObject, propertyPath, indexes, rank + 1);
                    i++;
                }
            }


        }

        private static PropertyHolder BuildHolder(IGrouping<string, Dictionary<string, object>> group)
        {
            PropertyHolder propertyHolder = new PropertyHolder();
            Dictionary<string, List<Dictionary<string, object>>> subDataTree = new Dictionary<string, List<Dictionary<string, object>>>();

            foreach (Dictionary<string, object> item in group)
            {
                foreach (KeyValuePair<string, object> item1 in item)
                {
                    if (item1.Value?.GetType() == typeof(Dictionary<string, object>))
                    {
                        Dictionary<string, object> nestedItem = (Dictionary<string, object>)item1.Value;
                        if (subDataTree.ContainsKey(item1.Key))
                        {
                            subDataTree[item1.Key].Add(nestedItem);
                        }
                        else
                        {
                            List<Dictionary<string, object>> nestedItems = new List<Dictionary<string, object>>();
                            nestedItems.Add(nestedItem);
                            subDataTree.Add(item1.Key, nestedItems);
                        }
                    }
                    else
                    {
                        if (!propertyHolder.Fields.ContainsKey(item1.Key))
                        {
                            propertyHolder.Fields.Add(item1.Key, item1.Value);
                        }
                    }
                }
            }

            foreach (KeyValuePair<string, List<Dictionary<string, object>>> propertySubDataTree in subDataTree)
            {
                IEnumerable<IGrouping<string, Dictionary<string, object>>> subGroups = propertySubDataTree.Value.GroupBy(d => GroupingString(d));

                List<PropertyHolder> subPropertiesHolders = new List<PropertyHolder>();
                foreach (IGrouping<string, Dictionary<string, object>> subGroup in subGroups)
                {
                    PropertyHolder nestedPropertyHolder = BuildHolder(subGroup);
                    subPropertiesHolders.Add(nestedPropertyHolder);
                }

                propertyHolder.Properties.Add(propertySubDataTree.Key, subPropertiesHolders);
            }


            return propertyHolder;
        }

        private static List<Dictionary<string, object>> CreateDataTree(object[,] inputArray, string[]? dateColumns = null)
        {

            string[] headers = Enumerable.Range(0, inputArray.GetLength(1))
                .Select(x => inputArray[0, x].ToString())
                .ToArray();

            bool[] isDateHeaders = Enumerable.Repeat(false, headers.Length).ToArray();

            if (dateColumns != null)
            {
                for (int j = 0; j < headers.Length; j++)
                {
                    if (!dateColumns.Contains(headers[j])) continue;
                    isDateHeaders[j] = true;
                }
            }


            List<Dictionary<string, object>> classInfoList = new List<Dictionary<string, object>>();

            for (int i = 1; i < inputArray.GetLength(0); i++)
            {
                object[] data = Enumerable.Range(0, inputArray.GetLength(1))
                .Select(x => inputArray[i, x])
                .ToArray();

                Dictionary<string, object> classInfo = new Dictionary<string, object>();

                for (int j = 0; j < headers.Length; j++)
                {
                    string header = headers[j];
                    bool isDateHeader = isDateHeaders[j];

                    object value = data[j];

                    if (isDateHeader)
                    {
                        if (value.GetType() == typeof(double))
                        {
                            value = DateTime.FromOADate((double)value);
                        }
                        else
                        {
                            value = null;
                        }
                    }

                    string[] nestedHeaders = header.Split('.');
                    if (nestedHeaders.Length > 1)
                    {
                        Dictionary<string, object> nestedDict = new Dictionary<string, object>();
                        Dictionary<string, object> currentDict = classInfo;

                        for (int k = 0; k < nestedHeaders.Length - 1; k++)
                        {
                            string nestedHeader = nestedHeaders[k];
                            if (!currentDict.ContainsKey(nestedHeader))
                            {
                                Dictionary<string, object> newDict = new Dictionary<string, object>();
                                currentDict[nestedHeader] = newDict;
                                currentDict = newDict;
                            }
                            else
                            {
                                currentDict = (Dictionary<string, object>)currentDict[nestedHeader];
                            }
                        }

                        currentDict[nestedHeaders.Last()] = value;
                    }
                    else
                    {
                        classInfo[header] = value;
                    }
                }

                classInfoList.Add(classInfo);
            }

            return classInfoList;
        }

        private static string GroupingString(Dictionary<string, object> dictionary)
        {
            string groupingString = "";
            foreach (KeyValuePair<string, object> keyValue in dictionary)
            {
                if (keyValue.Value?.GetType() != typeof(Dictionary<string, object>))
                {
                    string value = keyValue.Value?.ToString();

                    if (keyValue.Value == null)
                    {
                        value = "";
                    }
                    else if (keyValue.Value?.GetType().FullName == "ExcelDna.Integration.ExcelEmpty")
                    {
                        value = "";
                    }

                    groupingString = groupingString + ";" + value;
                }
            }

            return groupingString;
        }



        private static Dictionary<string, Type> ListColumns(object[,] inputArray, string[]? dateColumns = null)
        {
            // Build an dictonary of field
            Dictionary<string, Type> columnsWithType = new Dictionary<string, Type>();


            for (int i = 0; i < inputArray.GetLength(1); i++)
            {
                string name = inputArray[0, i].ToString();

                Type fieldType = GetNullableType(inputArray[1, i].GetType());

                if (dateColumns != null && dateColumns.Contains(name))
                {
                    fieldType = typeof(DateTime?);
                }
                else
                {
                    int rowIndex = 1;
                    while (fieldType.FullName == "ExcelDna.Integration.ExcelEmpty" && rowIndex < inputArray.GetLength(0))
                    {
                        fieldType = GetNullableType(inputArray[rowIndex, i].GetType());
                        rowIndex++;
                    }

                    if (fieldType.FullName == "ExcelDna.Integration.ExcelEmpty")
                    {
                        fieldType = typeof(string);
                    }
                }

                columnsWithType.Add(name, fieldType);
            }

            return columnsWithType;
        }

        private static void SetPropertyValue(object parentTarget, string compoundProperty, object value, int[] indexes, int rank)
        {
            string[] bits = compoundProperty.Split('.');
            int pathLenght = bits.Length;

            for (int i = 0; i < bits.Length - 1; i++)
            {
                if (IsList(parentTarget))
                {
                    IList list = parentTarget as IList;
                    int parentIndex = indexes[i];
                    parentTarget = list[parentIndex];
                }

                PropertyInfo propertyToGet = parentTarget.GetType().GetProperty(bits[i]);
                if (propertyToGet == null) { return; }
                object target = propertyToGet.GetValue(parentTarget, null);

                if (target == null)
                {
                    // Create a new list of object to be added to the parent object
                    target = Activator.CreateInstance(propertyToGet.PropertyType);
                    propertyToGet.SetValue(parentTarget, target);
                }

                parentTarget = target;
            }

            if (IsList(parentTarget))
            {
                // Does an object exist at the given index
                IList list = parentTarget as IList;
                if (list.Count <= indexes[rank])
                {
                    // Add a new object to the list
                    Type type = parentTarget.GetType().GetGenericArguments()[0];
                    object objTemp = Activator.CreateInstance(type);
                    PropertyInfo propertyToSet = objTemp.GetType().GetProperty(bits.Last());
                    if (value.GetType().FullName == "ExcelDna.Integration.ExcelEmpty")
                    {
                        propertyToSet.SetValue(objTemp, null, null);
                    }
                    else
                    {
                        propertyToSet.SetValue(objTemp, value, null);
                    }

                    parentTarget.GetType().GetMethod("Add").Invoke(parentTarget, new[] { objTemp });
                }
                else
                {
                    // Add the property to the given object in the list
                    object objTemp = list[indexes[rank]];
                    PropertyInfo propertyToSet = objTemp.GetType().GetProperty(bits.Last());

                    if (value.GetType().FullName == "ExcelDna.Integration.ExcelEmpty")
                    {
                        propertyToSet.SetValue(objTemp, null, null);
                    }
                    else
                    {
                        if (value.GetType() == typeof(string) && propertyToSet.PropertyType != typeof(string))
                        {
                            propertyToSet.SetValue(objTemp, null, null);
                        }
                        else
                        {
                            propertyToSet.SetValue(objTemp, value, null);
                        }
                    }
                }



            }
            else
            {
                PropertyInfo propertyToSet = parentTarget.GetType().GetProperty(bits.Last());

                if (value?.GetType().FullName == "ExcelDna.Integration.ExcelEmpty")
                {
                    propertyToSet.SetValue(parentTarget, null, null);
                }
                else if(value?.GetType() == typeof(string) && propertyToSet.PropertyType == typeof(double?))
                {
                    propertyToSet.SetValue(parentTarget, null, null);
                }
                else
                {
                    propertyToSet.SetValue(parentTarget, value, null);
                }
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

        public static bool IsList(object o)
        {
            if (o == null) return false;
            return o is IList &&
                   o.GetType().IsGenericType &&
                   o.GetType().GetGenericTypeDefinition().IsAssignableFrom(typeof(List<>));
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
                    Type genericListType = typeof(List<>).MakeGenericType(BuildType(nestedType, field.Key));
                    CreateProperty(tb, field.Key, genericListType);
                }
            }

            Type objectType = tb.CreateType();
            return objectType;
        }
    }

}
