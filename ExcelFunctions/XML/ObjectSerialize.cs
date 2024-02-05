using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace ExcelFunctions.XML
{
    [Serializable]
    public class ObjectSerialize : IXmlSerializable
    {
        public List<object> ObjectList { get; set; }

        public XmlSchema GetSchema()
        {
            return new XmlSchema();
        }

        public void ReadXml(XmlReader reader)
        {

        }

        public void WriteXml(XmlWriter writer)
        {
            string objectName = ObjectList[0].GetType().Name;
            foreach (object obj in ObjectList)
            {
                //Provide elements for object item
                writer.WriteStartElement(objectName);
                System.Reflection.PropertyInfo[] properties = obj.GetType().GetProperties();
                foreach (System.Reflection.PropertyInfo propertyInfo in properties)
                {
                    object value = propertyInfo.GetValue(obj);
                    string textValue = "";

                    if (value != null)
                    {
                        if (value.GetType() != typeof(ExcelDna.Integration.ExcelEmpty))
                        {
                            textValue = value?.ToString();
                        }
                    }


                    //Provide elements for per property
                    writer.WriteElementString(propertyInfo.Name, textValue);
                }
                writer.WriteEndElement();
            }
        }
    }
}


