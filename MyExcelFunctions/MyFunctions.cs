using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.IO.Compression;
using System.Reflection;
using System.Xml.Serialization;
using MyExcelFunctions.XML;
using System.Xml;
using System.Data;
using System.Xml.Linq;
using System.Drawing.Imaging;

namespace MyExcelFunctions
{
    public static class MyFunctions
    {

        [ExcelFunction(Category = "My functions", Description = "Insert a picture.", HelpTopic = "Insert a picture.")]
        public static object INSERTPICTURE(
    [ExcelArgument("path", Name = "path", Description = "A path to the picture.")] string path,
    [ExcelArgument("width", Name = "width", Description = "[optional] The width of the picture.")] string width,
    [ExcelArgument("height", Name = "height", Description = "[optional] The height of the picture.")] string height)
        {
            try
            {

                Excel.Application excelApplication = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Worksheet ws = (Excel.Worksheet)excelApplication.ActiveSheet;

                ExcelReference caller = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);

                Excel.Range oRange = (Excel.Range)ws.Cells[caller.RowFirst + 1, caller.ColumnFirst + 1];

                try
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("fr-FR");
                    //Get range value
                    int? value = null;

                    //System.Globalization.CultureInfo cultureInfo = new System.Globalization.CultureInfo("en-US");
                    //var test = oRange.GetType().InvokeMember("Value2", BindingFlags.InvokeMethod, null, oRange, null);

                    //var cellTest = ws.Cells[caller.RowFirst + 1, caller.ColumnFirst + 1].Values2;
                    var cellValue = (string)oRange.Value2;
                    value = Convert.ToInt16(cellValue);

                    if (value != null)
                    {
                        foreach (Excel.Shape shape in ws.Shapes)
                        {
                            if (shape.ID == oRange.Value)
                            {
                                shape.Delete();
                            }
                        }
                    }
                }
                catch
                {

                }



                float Left = (float)((double)oRange.Left);
                float Top = (float)((double)oRange.Top);

                System.Drawing.Image img = System.Drawing.Image.FromFile(path);

                double ratio = img.Width / img.Height;
                float pictureHeight = 50;
                float pictureWidth = (float)Math.Round(pictureHeight * ratio);

                if (height == "" && width == "")
                {

                }
                else if (height == "")
                {
                    pictureWidth = (float)Convert.ToInt16(width);
                    pictureHeight = (float)Math.Round(pictureWidth / ratio);
                }
                else if (width == "")
                {

                    pictureHeight = (float)Convert.ToInt16(height);
                    pictureWidth = (float)Math.Round(pictureHeight * ratio);
                }



                Excel.Shape picture = ws.Shapes.AddPicture(path, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, pictureWidth, pictureHeight);

                return picture.ID.ToString(); ;
            }
            catch (Exception ex)
            {
                string test = ex.Message;
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "My functions", Description = "Converts a given number into its full text in french.", HelpTopic = "Converts a given number into its full text in french.")]
        public static object CONVERTTOFRENCH(
    [ExcelArgument("input", Name = "input", Description = "The number to be spelled")] double value,
    [ExcelArgument("portion", Name = "portion", Description = "[optional] Type 0 for the interger portion, 1 for decimal portion, 2 for both, 3 for 2 digits. Default is 0")] int portion)
        {
            try
            {
                TexteEnLettre texteEnLettre = new TexteEnLettre();
                if (portion == 3)
                {
                    value = Math.Round(value, 2);
                }

                int wholePart = (int)Math.Truncate(value);
                decimal decimalValue = (decimal)(value - Math.Truncate(value));
                long decimalPart = 0;
                long twoDigitDecimalPart = 0;
                if (decimalValue != 0)
                {
                    string decimalString = value.ToString().Split(System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator.ToCharArray())[1];
                    if (decimalString.Length > 19)
                    {
                        decimalString = decimalString.Substring(0, 19);
                    }
                    string TwoDigitdecimalString = decimalString;

                    if (decimalString.Length == 0)
                    {
                        TwoDigitdecimalString = "0";
                    }
                    else if (decimalString.Length == 1)
                    {
                        TwoDigitdecimalString = decimalString + "0";
                    }
                    else
                    {
                        TwoDigitdecimalString = decimalString.Substring(0, 2);
                    }

                    decimalPart = Convert.ToInt64(decimalString);
                    twoDigitDecimalPart = Convert.ToInt64(TwoDigitdecimalString);
                }


                string texte = texteEnLettre.IntToFr(wholePart);

                if (portion == 0)
                {
                    texte = texteEnLettre.IntToFr(wholePart);
                }
                else if (portion == 1)
                {
                    texte = texteEnLettre.IntToFr(decimalPart);
                }
                else if (portion == 2)
                {
                    texte = texteEnLettre.IntToFr(wholePart) + " virgule " + texteEnLettre.IntToFr(decimalPart);
                }
                else if (portion == 3)
                {
                    texte = texteEnLettre.IntToFr(twoDigitDecimalPart);
                }

                //Clean text
                texte = texte.Replace("  ", " ");
                texte = texte.Trim(' ');
                // Ajoute une majuscule au début
                return texte.First().ToString().ToUpper() + texte.Substring(1);
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }


        [ExcelFunction(Category = "My functions", Description = "Create a table from sets of possible values.", HelpTopic = "Create a table from sets of possible values.")]
        public static object GENERATELISTFROMVALUES(
    [ExcelArgument("values", Name = "values", Description = "A series of values in columns.")] object values)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;
                    //Create a list of list of string
                    List<List<object>> columns = new List<List<object>>();

                    for (int i = 0; i < inputArray.GetLength(1); i++)
                    {
                        List<object> column = new List<object>();

                        for (int j = 0; j < inputArray.GetLength(0); j++)
                        {
                            if (inputArray[j, i] != ExcelDna.Integration.ExcelEmpty.Value)
                            {
                                column.Add(inputArray[j, i]);
                            }
                        }
                        columns.Add(column);
                    }

                    int lineNumber = columns.Aggregate(1, (x, y) => x * y.Count);
                    int columnNumber = columns.Count;

                    object[,] outputTable = new object[lineNumber, columnNumber];

                    //for (int i = 0; i < outputTable.GetLength(0); i++)
                    //{
                    //    for (int j = 0; j < outputTable.GetLength(1); j++)
                    //    {
                    //        if (inputArray[i, j] != ExcelDna.Integration.ExcelEmpty.Value)
                    //        {
                    //            outputTable[i, j] = 2;//columns[i][j];
                    //        }
                    //    }
                    //}

                    //List<List<object>> query = CrossJoinList(columns[0], columns[1]);

                    IEnumerable<IEnumerable<object>> query = CartesianProduct(columns);


                    int k = 0;
                    foreach (IEnumerable<object> objects in query)
                    {
                        int l = 0;
                        foreach (object obj in objects)
                        {
                            outputTable[k, l] = obj;
                            l++;
                        }
                        k++;
                    }
                    return outputTable;
                }
                catch
                {
                    return new object[,] { { ExcelDna.Integration.ExcelError.ExcelErrorNA } };
                }
            }
            else
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }

        }

        private static List<List<object>> CrossJoinList(List<object> list1, List<object> list2)
        {
            List<List<object>> query = list1.SelectMany(list0 => list2, (object0, object1) => new List<object>() { object0, object1 }).ToList<List<object>>();

            return query;
        }

        private static IEnumerable<IEnumerable<T>> CartesianProduct<T>(IEnumerable<IEnumerable<T>> sequences)
        {
            if (sequences == null)
            {
                return null;
            }

            IEnumerable<IEnumerable<T>> emptyProduct = new[] { Enumerable.Empty<T>() };

            return sequences.Aggregate(
                emptyProduct,
                (accumulator, sequence) => accumulator.SelectMany(
                    accseq => sequence,
                    (accseq, item) => accseq.Concat(new[] { item })));
        }


        [ExcelFunction(Category = "My functions", Description = "Get the number of UP for a givn number of occupants", HelpTopic = "Get the number of UP for a givn number of occupants")]
        public static object NOMBREUP(
    [ExcelArgument("rank", Name = "effectif", Description = "The number of occupants")] int effectif)
        {
            try
            {
                if (effectif < 20) { return 1; }
                else if (20 <= effectif && effectif <= 50) { return 2; } //it is actually more complex than that, I am being consevative here
                else if (50 < effectif && effectif <= 100) { return 2; }
                else if (100 < effectif && effectif <= 200) { return 3; }
                else if (200 < effectif && effectif <= 300) { return 4; }
                else if (300 < effectif && effectif <= 400) { return 5; }
                else if (400 < effectif && effectif <= 500) { return 6; }
                else if (effectif > 500)
                {
                    return (int)Math.Ceiling((double)effectif / 100);
                }
                else
                {
                    return 0;
                }
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "My functions", Description = "Get the number of evacuation paths for a given number of occupants", HelpTopic = "Get the number of evacuation paths for a given number of occupants")]
        public static object NOMBREDEGAGEMENTS(
[ExcelArgument("rank", Name = "effectif", Description = "The number of occupants")] int effectif)
        {
            try
            {
                if (effectif < 20) { return 1; }
                else if (20 <= effectif && effectif <= 50) { return 2; } //it is actually more complex than that, I am being consevative here
                else if (50 < effectif && effectif <= 100) { return 2; }
                else if (100 < effectif && effectif <= 200) { return 2; }
                else if (200 < effectif && effectif <= 300) { return 2; }
                else if (300 < effectif && effectif <= 400) { return 2; }
                else if (400 < effectif && effectif <= 500) { return 2; }
                else if (effectif > 500)
                {
                    return 2 + (int)Math.Ceiling((double)(effectif - 500) / 500);
                }
                else
                {
                    return 0;
                }
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "My functions", Description = "Get the number of parking slot available in a given lenght", HelpTopic = "Get the number of parking slot available in a given lenght")]
        public static object NOMBREDEPLACES(
            [ExcelArgument("lenght", Name = "lenght", Description = "The available lenght")] double lenght,
            [ExcelArgument("column", Name = "columnWidth", Description = "The width of a column")] double columnWidth)
        {
            try
            {

                if (lenght < 2.5)
                {
                    return 0;
                }
                else
                {
                    double cumulatedLenght = 0;
                    double nextPlaceLenght = 2.5; // 2.3, 2.5, 0.4
                    double blockLenght = 2.3 + 2.5 * 2 + columnWidth;
                    int placeNumber = 0;

                    // Count the number of blocks
                    int blockNumber = Convert.ToInt16(Math.Floor(lenght / blockLenght));
                    placeNumber = placeNumber + blockNumber * 3;

                    // Count the remaining places
                    double remainingLenght = Math.Round(lenght - blockNumber * blockLenght, 6);

                    if (remainingLenght < 2.5)
                    {
                        // I can't any place
                        placeNumber = placeNumber + 0;
                    }
                    else if (remainingLenght >= 2.5 && remainingLenght < 2.5 * 2)
                    {
                        // I can add one place
                        placeNumber = placeNumber + 1;
                    }
                    else if (remainingLenght >= 2 * 2.5 && remainingLenght < 2.5 * 2 + 2.3)
                    {
                        // I can add two places
                        placeNumber = placeNumber + 2;
                    }
                    else if (remainingLenght >= 2.5 * 2 + 2.3)
                    {
                        // I can add three places
                        placeNumber = placeNumber + 3;
                    }
                    else
                    {
                        // I can't any place
                        placeNumber = placeNumber + 0;
                    }

                    return placeNumber;

                }

            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

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

        [ExcelFunction(Category = "My functions", Description = "Replace string in text based on an array.", HelpTopic = "Replace string in text based on an array.")]
        public static object FINDANDREPLACE(
[ExcelArgument("table", Name = "table", Description = "The table where to look for values")] object values,
[ExcelArgument("column", Name = "column", Description = "The column where to find replacement values")] object column,
[ExcelArgument("text", Name = "text", Description = "The text to be replaced")] object text)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;
                    string textValue = (string)text;
                    int columnIndex = 0;
                    bool parseResult = int.TryParse(column.ToString(), out columnIndex);

                    if (!parseResult) return ExcelDna.Integration.ExcelError.ExcelErrorNA;

                    Dictionary<string, string> keyDictionary = new Dictionary<string, string>();

                    for (int i = 0; i < inputArray.GetLength(0); i++)
                    {
                        string key = inputArray[i, 0].ToString();
                        if (textValue.Contains(key))
                        {
                            string replacementValue = inputArray[i, columnIndex].ToString();
                            keyDictionary.Add(key, replacementValue);
                        }
                    }

                    string textReturn = textValue;
                    foreach (string key in keyDictionary.Keys)
                    {
                        textReturn = textReturn.Replace(key, keyDictionary[key]);
                    }

                    return textReturn;

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

        private static Type GetNullableType(Type type)
        {
            // Use Nullable.GetUnderlyingType() to remove the Nullable<T> wrapper if type is already nullable.
            type = Nullable.GetUnderlyingType(type) ?? type; // avoid type becoming null
            if (type.IsValueType)
                return typeof(Nullable<>).MakeGenericType(type);
            else
                return type;
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
            Type type = XML.XmlTypeBuilder.CompileResultType(fields.ToArray(), rowName.ToString());

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

        [ExcelFunction(Category = "My functions", Description = "Create a table from sets of possible values.", HelpTopic = "Create a table from sets of possible values.")]
        public static object BIMSYNCFOLDERS(
    [ExcelArgument("folders", Name = "folders", Description = "Two columns to describe the folders")] object values)
        {
            if (values is object[,])
            {
                try
                {
                    object[,] inputArray = (object[,])values;

                    // Create a list of folder
                    List<MyFolder> myFolders = new List<MyFolder>();

                    for (int i = 0; i < inputArray.GetLength(0); i++)
                    {
                        // Check if the cell is not empty
                        if (inputArray[i, 0] != ExcelDna.Integration.ExcelEmpty.Value)
                        {
                            string parentId = "";
                            if (inputArray[i, 1] != ExcelDna.Integration.ExcelEmpty.Value) { parentId = (string)inputArray[i, 1]; }
                            MyFolder myFolder = new MyFolder(parentId, (string)inputArray[i, 0]);
                            myFolders.Add(myFolder);
                        }
                    }

                    Dictionary<string, BimsyncFolder> lookup = new Dictionary<string, BimsyncFolder>();
                    myFolders.ForEach(x => lookup.Add(x.ID, new BimsyncFolder { AssociatedFolder = x }));

                    foreach (var item in lookup.Values)
                    {
                        BimsyncFolder proposedParent;
                        if (lookup.TryGetValue(item.AssociatedFolder.ParentID, out proposedParent))
                        {
                            item.Parent = proposedParent;
                            proposedParent.Children.Add(item);
                        }
                    }
                    List<BimsyncFolder> bimsyncFolders = lookup.Values.Where(x => x.Parent == null).ToList();

                    List<Folder> folders = new List<Folder>();

                    foreach (BimsyncFolder bimsyncFolder in bimsyncFolders)
                    {
                        folders.Add(new Folder(bimsyncFolder));
                    }

                    JsonSerializerSettings jsonSerializerSettings = new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore
                    };

                    string json = JsonConvert.SerializeObject(folders, jsonSerializerSettings);

                    return "[{\"folders\": " + json + "}]";

                }
                catch
                {
                    return new object[,] { { ExcelDna.Integration.ExcelError.ExcelErrorNA } };
                }
            }
            else
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }

        }

        [ExcelFunction(Category = "My functions", Description = "Escape non ASCII characters with their unicode value", HelpTopic = "Escape non ASCII characters with their unicode value")]
        public static object ENCODENONASCIICHARACTERS(
            [ExcelArgument("input", Name = "input", Description = "The string to escape.")] string input)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                foreach (char c in input)
                {
                    if (c > 127)
                    {
                        // This character is too big for ASCII
                        string encodedValue = "\\u" + ((int)c).ToString("x4");
                        sb.Append(encodedValue);
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
                return sb.ToString();
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        [ExcelFunction(Category = "My functions", Description = "Télécharge une quantité d'une fiche produit de la base Inies", HelpTopic = "Télécharge une quantité d'une fiche produit de la base Inies")]
        public static object INIESQUANTITE(
    [ExcelArgument("Norme", Name = "Norme", Description = "Le numéro de la norme de la quantité recherché")] int norme,
    [ExcelArgument("Phase", Name = "Phase", Description = "Le numéro de la phase de la quantité recherché")] int phase,
    [ExcelArgument("NumFiche", Name = "NumFiche", Description = "Le numéro de la fiche produit.")] int NumFiche)
        {
            // Don't do anything else here - might run at unexpected times...
            return ExcelAsyncUtil.Run("INIES", new object[] { NumFiche, norme, phase },
                delegate { return GetQuantite(NumFiche, norme, phase); });

        }

        [ExcelFunction(Category = "My functions", Description = "Retourne le lien vers la fiche produit INIES", HelpTopic = "Retourne le lien vers la fiche produit INIES")]
        public static object INIESLINK(
[ExcelArgument("NumFiche", Name = "NumFiche", Description = "Le numéro de la fiche produit.")] int NumFiche)
        {
            try
            {
                return $"https://www.base-inies.fr/iniesV4/dist/consultation.html?id={NumFiche}";
            }
            catch
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        private static object GetQuantite(int NumFiche, int normeId, int phaseId)
        {
            try
            {
                Produit produit = GetProduit(NumFiche).Result;

                if (produit == null) return ExcelDna.Integration.ExcelError.ExcelErrorNull;

                TINDICATEURQUANTITE indicateursQuantité = produit.T_INDICATEUR_QUANTITEs.Where(i => i.ID_INDICATEUR_NORME == normeId && i.ID_PHASE_NORME == phaseId).FirstOrDefault();

                if (indicateursQuantité != null)
                {
                    return indicateursQuantité.QUANTITE;
                }

                return ExcelDna.Integration.ExcelError.ExcelErrorNull;
            }
            catch (Exception ex)
            {
                return ExcelDna.Integration.ExcelError.ExcelErrorNA;
            }
        }

        private static async Task<Produit> GetProduit(int NumFiche)
        {
            ICredentials credentials = CredentialCache.DefaultCredentials;
            IWebProxy proxy = WebRequest.DefaultWebProxy;
            proxy.Credentials = credentials;

            HttpClientHandler httpClientHandler = new HttpClientHandler()
            {
                Proxy = proxy,
                AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip
            };

            using (var client = new HttpClient(httpClientHandler))
            {

                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
                client.DefaultRequestHeaders.AcceptLanguage.Add(new StringWithQualityHeaderValue("en-US"));
                ////client.DefaultRequestHeaders.AcceptLanguage.Add(new StringWithQualityHeaderValue("en;q=0.9"));
                ////client.DefaultRequestHeaders.AcceptLanguage.Add(new StringWithQualityHeaderValue("fr;q=0.8"));
                //client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("ScraperBot", "1.0"));

                HttpResponseMessage httpResponseMessage = await client.GetAsync($"https://www.base-inies.fr/iniesV4/dist/api/Produit/{NumFiche}");

                using (Stream csStream = new GZipStream(await httpResponseMessage.Content.ReadAsStreamAsync(), CompressionMode.Decompress))
                {
                    // convert stream to string
                    StreamReader reader = new StreamReader(csStream);
                    string responseString = reader.ReadToEnd();

                    Produit produit = JsonConvert.DeserializeObject<Produit>(responseString);

                    return produit;

                }

            }
        }


        [ExcelFunction(Category = "My functions", Description = "Télécharge une quantité d'une fiche produit de la base Inies", HelpTopic = "Télécharge une quantité d'une fiche produit de la base Inies")]
        public static object INIESUNITE(
[ExcelArgument("Norme", Name = "Norme", Description = "Le numéro de la norme de la quantité recherché")] int normeId)
        {
            List<NORME> normes = ReadConfigurations<List<NORME>>("normes.json");
            List<UNITE> unites = ReadConfigurations<List<UNITE>>("unite.json");
            NORME norme = normes.Where(n => n.ID_NORME == 2).FirstOrDefault();
            TINDICATEURNORME tINDICATEURNORME = norme.T_INDICATEUR_NORMEs.Where(i => i.ID_INDICATEUR_NORME == normeId).FirstOrDefault();
            UNITE unite = unites.Where(u => u.ID_UNITE == tINDICATEURNORME.T_INDICATEURs.ID_UNITE).FirstOrDefault();
            return unite.NOM_UNITE;

        }

        private static T ReadConfigurations<T>(string fileName)
        {

            //or from the entry point to the application - there is a difference!
            string[] names = Assembly.GetExecutingAssembly().GetManifestResourceNames();


            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream("MyExcelFunctions.Resources." + fileName))
            {
                // convert stream to string
                StreamReader reader = new StreamReader(stream);
                string responseString = reader.ReadToEnd();
                T value = JsonConvert.DeserializeObject<T>(responseString);
                return value;
            }
        }
    }
}
