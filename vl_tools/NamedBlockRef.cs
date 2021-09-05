using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.DatabaseServices;

namespace vl_tools
{
    public class NamedBlockRef
    {
        public string Name;
        public BlockReference bRef;
        public NamedBlockRef(BlockReference blockRef, string sname)
        {
            bRef = blockRef;
            Name = sname;            
        }
    }

    public class NamedBlockRefsCollection : Collection<NamedBlockRef>
    {
        
        protected override void InsertItem(int index, NamedBlockRef item)
        {
            if (this.Count > 0)
            {
                if (item.Name != this.Items[0].Name) throw new ArgumentException("Другое имя при попытке добавить в коллекцию");
            }
            base.InsertItem(index, item);
        }

        protected override void SetItem(int index, NamedBlockRef item)
        {
            if (this.Count > 0)
            {
                if (item.Name != this.Items[0].Name) throw new ArgumentException("Другое имя при попытке добавить в коллекцию");
            }
            base.SetItem(index, item);
        }

        public System.Data.DataTable GetAttsTable()
        {
            try
            {
                List<string> Tags = new List<string>();
                System.Data.DataTable table = new System.Data.DataTable();
                Database dbCurrent = HostApplicationServices.WorkingDatabase;
                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    //create columns
                    foreach (NamedBlockRef nref in this.Items)
                    {
                        AttributeCollection attcol = nref.bRef.AttributeCollection;
                        foreach (ObjectId att in attcol)
                        {
                            AttributeReference atRef = (AttributeReference)tr.GetObject(att, OpenMode.ForRead);
                            string tag = atRef.Tag;
                            if (Tags.Contains(tag))
                            {
                            }
                            else
                            {
                                table.Columns.Add(tag);
                                Tags.Add(tag);
                            }
                        }
                    }
                    //fill rows
                    foreach (NamedBlockRef nref in this.Items)
                    {
                        System.Data.DataRow row = table.NewRow();
                        AttributeCollection attcol = nref.bRef.AttributeCollection;
                        foreach (ObjectId att in attcol)
                        {
                            AttributeReference atRef = (AttributeReference)tr.GetObject(att, OpenMode.ForRead);
                            row[atRef.Tag] = atRef.TextString;
                        }
                        table.Rows.Add(row);
                    }
                    tr.Commit();
                }
                return table;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void UpdateAtts()
        {
        }
    }
}
