namespace LEP
{
    using Autodesk.AutoCAD.ApplicationServices.Core;
    using Autodesk.AutoCAD.DatabaseServices;
    using System;
    using System.Data;
    using System.IO;
    using System.Runtime.CompilerServices;
    using System.Text;
    using System.Xml.Linq;
    using System.Xml.Serialization;

    [Serializable]
    public class PlineObject : DwgObject
    {
        public PlineObject()
        {
            this.Length_dwg = 0.0;
            this.Length_redef = 0.0;
            this.UseRedefLength = false;
            this.AdditionalPercent = 0.0;
            this.Multiplier = 1;
            this.AdditionalLength1 = 0.0;
            this.AdditionalLength2 = 0.0;
        }

        //public PlineObject(DataRow dataRow)
        //{
        //    base.ObjId = ObjectId.Null;
        //    base.Name = dataRow.Field<string>("id_element");
        //    base.ObjectType = DAL.myDataSet.Tables["conductors"].Rows.Find(base.Name).Field<string>("id_type");
        //    base.Number = "0";
        //    base.HasExtData = false;
        //    this.Length_redef = 0.0;
        //    this.UseRedefLength = false;
        //    this.AdditionalPercent = DAL.preferences.AdditionalPercent;
        //    this.Multiplier = DAL.myDataSet.Tables["conductors"].Rows.Find(base.Name).Field<int>("multiplier");
        //    this.AdditionalLength1 = 0.0;
        //    this.AdditionalLength2 = 0.0;
        //    this.Start = string.Empty;
        //    this.End = string.Empty;
        //}

        public static double GetPolylineLength(ObjectId id)
        {
            using (Transaction transaction = Application.DocumentManager.MdiActiveDocument.Editor.Document.Database.TransactionManager.StartTransaction())
            {
                Entity entity = transaction.GetObject(id, OpenMode.ForRead) as Entity;
                if (entity.GetType() == typeof(Polyline))
                {
                    Polyline polyline = entity as Polyline;
                    if (polyline != null)
                    {
                        return polyline.Length;
                    }
                    return 0.0;
                }
                return 0.0;
            }
        }

        public static PlineObject Open(ObjectId id)
        {
            XElement xMLfromCADEntity = DwgObject.GetXMLfromCADEntity(id);
            if (xMLfromCADEntity == null)
            {
                return new PlineObject { HasExtData = false, ObjId = id };
            }
            XmlSerializer serializer = new XmlSerializer(typeof(PlineObject));
            PlineObject obj3 = (PlineObject) serializer.Deserialize(xMLfromCADEntity.CreateReader());
            obj3.HasExtData = true;
            obj3.ObjId = id;
            obj3.Length_dwg = GetPolylineLength(id);
            return obj3;
        }

        public XElement ToXElement()
        {
            XElement element;
            using (MemoryStream stream = new MemoryStream())
            {
                using (TextWriter writer = new StreamWriter(stream))
                {
                    new XmlSerializer(typeof(PlineObject)).Serialize(writer, this);
                    element = XElement.Parse(Encoding.UTF8.GetString(stream.ToArray()));
                }
            }
            return element;
        }

        public XElement ToXSpecification()
        {
            return new XElement("Element", new object[] { new XAttribute("type", base.ObjectType), new XElement("Specification", new object[] { new XAttribute("name", "FirstTable"), new XElement("Item", new object[] { new XAttribute("name", base.Name), new XAttribute("count", this.Length), new XAttribute("sendInSpecification", true), new XAttribute("comment", "") }) }) });
        }

        public double AdditionalLength1 { get; set; }

        public double AdditionalLength2 { get; set; }

        public double AdditionalPercent { get; set; }

        public string CalculationText
        {
            get
            {
                string str2;
                if (this.UseRedefLength)
                {
                    str2 = this.Length_redef.ToString();
                }
                else
                {
                    str2 = Math.Round(this.Length_dwg, 2).ToString();
                }
                string str3 = (1.0 + (this.AdditionalPercent / 100.0)).ToString();
                return ("L = " + str2 + " x " + this.Multiplier.ToString() + " x " + str3 + " + " + this.AdditionalLength1.ToString() + " + " + this.AdditionalLength2.ToString() + " = " + this.Length.ToString() + " м;");
            }
        }

        public string Comment { get; set; }

        public string End { get; set; }

        public double Length
        {
            get
            {
                double num;
                if (this.UseRedefLength)
                {
                    num = this.Length_redef;
                }
                else
                {
                    num = this.Length_dwg;
                }
                return Math.Ceiling((double) ((((num * this.Multiplier) * (1.0 + (this.AdditionalPercent / 100.0))) + this.AdditionalLength1) + this.AdditionalLength2));
            }
        }

        public double Length_dwg { get; set; }

        public double Length_redef { get; set; }

        public int Multiplier { get; set; }

        public string Start { get; set; }

        public bool UseRedefLength { get; set; }
    }
}

