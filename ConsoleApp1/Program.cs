using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("Hello");
            VLVolumeObject obj = new VLVolumeObject(1, "fer1", "job1", 2300.12, "km", "=3*2");
            //using (StringWriter textWriter = new StringWriter())
            //{
            //    XmlSerializer xmlSerializer = new XmlSerializer(typeof(VLVolumeObject));
            //    xmlSerializer.Serialize(textWriter, obj);
            //    Console.Write(textWriter.ToString());
            //}

            //VLVolumeCollection col = new VLVolumeCollection();
            //for (int i = 0; i < 10; i++) col.volumes.Add(obj);           
            //using (StringWriter textWriter = new StringWriter())
            //{
            //    XmlSerializer xmlSerializer = new XmlSerializer(typeof(VLVolumeCollection));
            //    xmlSerializer.Serialize(textWriter, col);
            //    Console.Write(textWriter.ToString());
            //}


            //XmlSerializer formatter = new XmlSerializer(typeof(VLVolumeCollection));
            //// получаем поток, куда будем записывать сериализованный объект
            //using (FileStream fs = new FileStream("volumes.xml", FileMode.OpenOrCreate))
            //{
            //    formatter.Serialize(fs, col);               
            //}


            //BinaryFormatter formatterb = new BinaryFormatter();
            //// получаем поток, куда будем записывать сериализованный объект
            //using (FileStream fs = new FileStream("volumes.dat", FileMode.OpenOrCreate))
            //{
            //    formatterb.Serialize(fs, col);
            //}
            DataSet set = new DataSet("ds");
            DataTable tbl = new DataTable("volumes");
            
            tbl.Columns.Add(new System.Data.DataColumn("id", typeof(int), "", MappingType.Attribute));
            tbl.Columns.Add(new System.Data.DataColumn("code", typeof(string), "", MappingType.Attribute));
            tbl.Columns.Add(new System.Data.DataColumn("name", typeof(string), "", MappingType.Attribute));
            tbl.Columns.Add(new System.Data.DataColumn("price", typeof(double), "", MappingType.Attribute));
            tbl.Columns.Add(new System.Data.DataColumn("unit", typeof(string), "", MappingType.Attribute));
            tbl.Columns.Add(new System.Data.DataColumn("formula", typeof(string), "", MappingType.Attribute));            
            for (int i = 0; i < 10; i++) tbl.Rows.Add(1, "fer1", "job1", 2300.12, "km", "=3*2");

            set.Tables.Add(tbl);
            using (StringWriter textWriter = new StringWriter())
            {
                set.WriteXml(textWriter);
                Console.Write(textWriter.ToString());
            }

            //// десериализация
            //using (FileStream fs = new FileStream("persons.xml", FileMode.OpenOrCreate))
            //{
            //    VLVolumeObject obj2 = (VLVolumeObject)formatter.Deserialize(fs);

            //    Console.WriteLine("Объект десериализован");                
            //}
            Console.ReadKey();
        }
    }
}
