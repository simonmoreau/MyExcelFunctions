﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace MyExcelFunctions.XML
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
            foreach (var obj in ObjectList)
            {
                //Provide elements for object item
                writer.WriteStartElement(objectName);
                var properties = obj.GetType().GetProperties();
                foreach (var propertyInfo in properties)
                {
                    //Provide elements for per property
                    writer.WriteElementString(propertyInfo.Name, propertyInfo.GetValue(obj).ToString());
                }
                writer.WriteEndElement();
            }
        }
    }
}


