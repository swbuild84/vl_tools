namespace LEP
{
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.ApplicationServices.Core;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.Runtime;
    //using LEP.Properties;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using System.Windows.Forms;
    using System.Xml.Linq;

    [Serializable]
    public class BlockObject : DwgObject
    {
        public BlockObject()
        {
        }

        public BlockObject(ObjectId id)
            : this(DwgObject.GetXMLfromCADEntity(id))
        {
            base.ObjId = id;
            Document mdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Editor editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
            using (mdiActiveDocument.LockDocument())
            {
                using (Transaction transaction = editor.Document.Database.TransactionManager.StartTransaction())
                {
                    BlockReference reference = (BlockReference)transaction.GetObject(base.ObjId, OpenMode.ForRead);
                    foreach (ObjectId id2 in reference.AttributeCollection)
                    {
                        AttributeReference reference2 = (AttributeReference)transaction.GetObject(id2, OpenMode.ForRead);
                        if (reference2.Tag == "SL_NUM")
                        {
                            base.Number = reference2.TextString.Trim();
                        }
                        if (reference2.Tag == "SL_NAME")
                        {
                            base.Name = reference2.TextString.Trim();
                        }
                    }
                    this.BlockName = reference.Name;
                }
            }
        }

        public BlockObject(XElement xData)
        {
            if (xData == null)
            {
                base.HasExtData = false;
            }
            else
            {
                base.Name = xData.Attribute("name").Value;
                base.ObjectType = xData.Attribute("type").Value;
                base.Number = xData.Attribute("number").Value;
                this.BlockName = string.Empty;
                base.HasExtData = true;
                //DataRow row = DAL.myDataSet.Tables["types_of_dwgobjects"].Rows.Find(base.ObjectType);
                this.Table_1 = DwgObject.CreateDataTable("firsttablename");
                this.Table_2 = DwgObject.CreateDataTable("secondtablename");
                foreach (XElement element2 in (from i in xData.Descendants("Specification")
                                               where i.Attribute("name").Value == "FirstTable"
                                               select i).First<XElement>().Elements("Item"))
                {
                    this.Table_1.Rows.Add(new object[] { element2.Attribute("name").Value, Convert.ToDouble(element2.Attribute("count").Value, CultureInfo.GetCultureInfo("en-US")), element2.Attribute("sendInSpecification").Value, element2.Attribute("comment").Value });
                }
                foreach (XElement element4 in (from i in xData.Descendants("Specification")
                                               where i.Attribute("name").Value == "SecondTable"
                                               select i).First<XElement>().Elements("Item"))
                {
                    this.Table_2.Rows.Add(new object[] { element4.Attribute("name").Value, Convert.ToDouble(element4.Attribute("count").Value, CultureInfo.GetCultureInfo("en-US")), element4.Attribute("sendInSpecification").Value, element4.Attribute("comment").Value });
                }
            }
        }

        public void SetAttributes()
        {
            if (base.ObjId.IsNull)
            {
                throw new ArgumentNullException("ObjId");
            }
            Document mdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Editor editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
            using (mdiActiveDocument.LockDocument())
            {
                using (Transaction transaction = editor.Document.Database.TransactionManager.StartTransaction())
                {
                    try
                    {
                        BlockReference reference = (BlockReference)transaction.GetObject(base.ObjId, OpenMode.ForRead);
                        foreach (ObjectId id in reference.AttributeCollection)
                        {
                            AttributeReference reference2 = (AttributeReference)transaction.GetObject(id, OpenMode.ForRead);
                            if (reference2.Tag == "SL_NUM")
                            {
                                reference2.UpgradeOpen();
                                reference2.TextString = base.Number;
                            }
                            if (reference2.Tag == "SL_NAME")
                            {
                                reference2.UpgradeOpen();
                                reference2.TextString = base.Name;
                            }
                        }
                        transaction.Commit();
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception)
                    {
                        throw;
                    }
                }
            }
        }

        public XElement ToXElement()
        {
            XElement element = new XElement("Element", new object[] { new XAttribute("type", base.ObjectType), new XAttribute("name", base.Name), new XAttribute("number", base.Number), new XAttribute("blockName", this.BlockName), new XAttribute("program", "ESMT_LEP_v1.0") });
            XElement content = new XElement("Specification", new XAttribute("name", "FirstTable"));
            XElement element3 = new XElement("Specification", new XAttribute("name", "SecondTable"));
            foreach (DataRow row in this.Table_1.Rows)
            {
                if ((row.RowState != DataRowState.Deleted) && (row.RowState != DataRowState.Detached))
                {
                    string str = row.Field<string>("comment");
                    if (str == null)
                    {
                        str = string.Empty;
                    }
                    content.Add(new XElement("Item", new object[] { new XAttribute("name", row.Field<string>("item_name")), new XAttribute("count", row.Field<double>("item_count").ToString(new CultureInfo("en-US"))), new XAttribute("sendInSpecification", row.Field<bool>("sendInSpecification")), new XAttribute("comment", str) }));
                }
            }
            foreach (DataRow row2 in this.Table_2.Rows)
            {
                if ((row2.RowState != DataRowState.Deleted) && (row2.RowState != DataRowState.Detached))
                {
                    string str2 = row2.Field<string>("comment");
                    if (str2 == null)
                    {
                        str2 = string.Empty;
                    }
                    element3.Add(new XElement("Item", new object[] { new XAttribute("name", row2.Field<string>("item_name")), new XAttribute("count", row2.Field<double>("item_count").ToString(new CultureInfo("en-US"))), new XAttribute("sendInSpecification", row2.Field<bool>("sendInSpecification")), new XAttribute("comment", str2) }));
                }
            }
            element.Add(content);
            element.Add(element3);
            return element;
        }

        public string BlockName { get; set; }

        public System.Data.DataTable Table_1 { get; set; }

        public System.Data.DataTable Table_2 { get; set; }
    }
}

