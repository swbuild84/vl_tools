namespace LEP
{
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.ApplicationServices.Core;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Runtime;
    using System;
    using System.Data;
    using System.Runtime.CompilerServices;
    using System.Xml.Linq;

    public class DwgObject
    {
        public DwgObject()
        {
        }

        public DwgObject(ObjectId id)
        {
            this.ObjId = id;
            XElement xMLfromCADEntity = GetXMLfromCADEntity(id);
            if (xMLfromCADEntity == null)
            {
                this.HasExtData = false;
            }
            else
            {
                //this.Name = xMLfromCADEntity.Element("Name").Value;
                //this.ObjectType = xMLfromCADEntity.Element("ObjectType").Value;
                //this.Number = xMLfromCADEntity.Element("Number").Value;
                this.HasExtData = true;
            }
        }

        public static System.Data.DataTable CreateDataTable(string nameOfTable)
        {
            System.Data.DataColumn[] columnArray = new System.Data.DataColumn[] { new System.Data.DataColumn("item_name", Type.GetType("System.String")), new System.Data.DataColumn("item_count", Type.GetType("System.Double")), new System.Data.DataColumn("sendInSpecification", Type.GetType("System.Boolean")), new System.Data.DataColumn("comment", Type.GetType("System.String")) };
            columnArray[2].DefaultValue = true;
            System.Data.DataTable table = new System.Data.DataTable(nameOfTable);
            foreach (System.Data.DataColumn column in columnArray)
            {
                table.Columns.Add(column);
            }
            table.PrimaryKey = new System.Data.DataColumn[] { table.Columns[0] };
            return table;
        }

        public static XElement GetXMLfromCADEntity(ObjectId id)
        {
            XElement element;
            Transaction transaction = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.Document.Database.TransactionManager.StartTransaction();
            try
            {
                Entity entity = (Entity)transaction.GetObject(id, OpenMode.ForRead);
                if (entity.ExtensionDictionary.IsNull)
                {
                    transaction.Commit();
                    return null;
                }
                DBDictionary dictionary = (DBDictionary)transaction.GetObject(entity.ExtensionDictionary, OpenMode.ForRead);
                string entryName = string.Empty;
                if (dictionary.Contains("ESMT_LEP_v1.0"))
                {
                    entryName = "ESMT_LEP_v1.0";
                }
                else
                {
                    entryName = string.Empty;
                }
                if (entryName == string.Empty)
                {
                    transaction.Commit();
                    return null;
                }
                ObjectId at = dictionary.GetAt(entryName);
                Xrecord xrecord = null;
                xrecord = (Xrecord)transaction.GetObject(at, OpenMode.ForRead);
                TypedValue[] valueArray = xrecord.Data.AsArray();
                transaction.Commit();
                element = XElement.Parse(valueArray[0].Value.ToString());
            }
            catch (Autodesk.AutoCAD.Runtime.Exception)
            {
                throw;
            }
            finally
            {
                transaction.Dispose();
            }
            return element;
        }

        public void SaveXMLtoCADEntity(XElement xData)
        {
            Document mdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            Editor editor = mdiActiveDocument.Editor;
            using (mdiActiveDocument.LockDocument())
            {
                using (Transaction transaction = editor.Document.Database.TransactionManager.StartTransaction())
                {
                    try
                    {
                        Entity entity = (Entity)transaction.GetObject(this.ObjId, OpenMode.ForWrite, false, true);
                        if (entity.ExtensionDictionary.IsNull)
                        {
                            entity.CreateExtensionDictionary();
                        }
                        DBDictionary dictionary = (DBDictionary)transaction.GetObject(entity.ExtensionDictionary, OpenMode.ForWrite, false, true);
                        if (dictionary.Contains("ESMT_LEP_v1.0"))
                        {
                            dictionary.Remove("ESMT_LEP_v1.0");
                        }
                        Xrecord newValue = new Xrecord();
                        TypedValue[] values = new TypedValue[] { new TypedValue(1, xData.ToString()) };
                        newValue.Data = new ResultBuffer(values);
                        dictionary.SetAt("ESMT_LEP_v1.0", newValue);
                        transaction.AddNewlyCreatedDBObject(newValue, true);
                        transaction.Commit();
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception)
                    {
                        throw;
                    }
                }
            }
        }

        public bool HasExtData { get; set; }

        public string Name { get; set; }

        public string Number { get; set; }

        public string ObjectType { get; set; }

        public ObjectId ObjId { get; set; }
    }
}

