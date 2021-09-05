using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace vl_tools
{
    [Serializable]
    public class VLBlockObj: VLDwgObject
    {
        //public string BlockName { get; set; }
        public System.Data.DataTable VolumesTable { get; set; }
        public VLBlockObj()
        {
            System.Data.DataColumn[] columnArray = new System.Data.DataColumn[] 
            { new System.Data.DataColumn("id", Type.GetType("System.Int64"),"", MappingType.Attribute),
                new System.Data.DataColumn("code", Type.GetType("System.String"),"", MappingType.Attribute),
                new System.Data.DataColumn("name", Type.GetType("System.String"),"", MappingType.Attribute),
                new System.Data.DataColumn("price", Type.GetType("System.String"),"", MappingType.Attribute),
                new System.Data.DataColumn("unit", Type.GetType("System.String"),"", MappingType.Attribute),
                new System.Data.DataColumn("count", Type.GetType("System.Double"),"", MappingType.Attribute),
                new System.Data.DataColumn("formula", Type.GetType("System.String"),"", MappingType.Attribute),
            };

            VolumesTable = new System.Data.DataTable("volumes");
            foreach (System.Data.DataColumn column in columnArray)
            {
                VolumesTable.Columns.Add(column);
            }
            //table.PrimaryKey = new System.Data.DataColumn[] { table.Columns[0] };
        }

        public static VLBlockObj Open(ObjectId id)
        {
            XElement xMLfromCADEntity = VLDwgObject.GetXMLfromCADEntity(id);
            if(xMLfromCADEntity == null)
            {
                return new VLBlockObj { HasExtData = false, ObjId = id };
            }

            System.Data.DataTable tbl = new System.Data.DataTable();
            //tbl.ReadXml(xMLfromCADEntity.);
            //VLBlockObj obj = new VLBlockObj();
            //obj.VolumesTable = tbl;
            XmlSerializer serializer = new XmlSerializer(typeof(VLBlockObj));
            VLBlockObj obj = (VLBlockObj)serializer.Deserialize(xMLfromCADEntity.CreateReader());
            obj.HasExtData = true;
            obj.ObjId = id;            
            return obj;
        }

        public bool TryToSave()
        {
            try
            {
                this.SaveXMLtoCADEntity(this.ToXElement());
                return true;
            }
            catch (Exception)
            {
                throw;
            }
            
        }

        public XElement ToXElement()
        {
            XElement element;
            using (MemoryStream stream = new MemoryStream())
            {
                using (TextWriter writer = new StreamWriter(stream))
                {
                    //this.VolumesTable.WriteXml(writer);
                    //element = XElement.Parse(Encoding.UTF8.GetString(stream.ToArray()));
                    new XmlSerializer(typeof(VLBlockObj)).Serialize(writer, this);
                    element = XElement.Parse(Encoding.UTF8.GetString(stream.ToArray()));
                }
            }
            return element;
        }
    }
}
