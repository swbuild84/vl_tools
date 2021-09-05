using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace vl_tools
{    
    public class VLDwgObject
    {
        [XmlIgnoreAttribute] public bool HasExtData { get; set; }
        [XmlIgnoreAttribute] public string ObjectType { get; set; }
        [XmlIgnoreAttribute] public ObjectId ObjId { get; set; }
        private XElement xMLfromCADEntity;

        public VLDwgObject()
        {
        }

        public VLDwgObject(ObjectId id)
        {
            this.ObjId = id;
            xMLfromCADEntity = GetXMLfromCADEntity(id);
            if (xMLfromCADEntity == null)
            {
                this.HasExtData = false;
            }
            else
            {               
                this.HasExtData = true;
            }
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
                if (dictionary.Contains("vl_tools_v1.0"))
                {
                    entryName = "vl_tools_v1.0";
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
                        if (dictionary.Contains("vl_tools_v1.0"))
                        {
                            dictionary.Remove("vl_tools_v1.0");
                        }
                        Xrecord newValue = new Xrecord();
                        TypedValue[] values = new TypedValue[] { new TypedValue(1, xData.ToString()) };
                        newValue.Data = new ResultBuffer(values);
                        dictionary.SetAt("vl_tools_v1.0", newValue);
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
    }
}
