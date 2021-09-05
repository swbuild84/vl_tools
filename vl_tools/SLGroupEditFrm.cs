using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using LEP;
using System.Globalization;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;


namespace vl_tools
{
    public partial class SLGroupEditFrm : Form
    {
        //private bool m_modified = false;    //флажок изменений
        ObjectId curRefid = ObjectId.Null;

        private BindingSource bindingSource1 = new BindingSource();        

        private System.Data.DataTable _tblNames;       
        private System.Data.DataTable _tblDetailsFirst; 
        private System.Data.DataTable _tblDetailsSecond;
        private System.Data.DataTable _unionTable;

        private List<BlockObject> _dwgObjcts;
        private List<bool> _modified=new List<bool>();

       

        public SLGroupEditFrm(List<BlockObject> dwgObjcts)
        {
            try
            {
                _dwgObjcts = dwgObjcts;
                ReadData();
                InitializeComponent();
            }
            catch (Exception ex)
            {                
                MessageBox.Show(ex.ToString());
            }
        }        

        private void ReadData()
        {
            List<string> detailsFirstTable = new List<string>(); //список попавшихся деталей
            List<string> detailsSecondTable = new List<string>(); //список попавшихся деталей            
            _tblNames = new System.Data.DataTable();    //таблица данных
            _tblDetailsFirst = new System.Data.DataTable();    //таблица данных
            _tblDetailsSecond = new System.Data.DataTable();    //таблица данных

            _tblNames.Columns.Add("НАИМЕНОВАНИЕ");
            _tblNames.Rows.Add("НОМЕР");
            _tblNames.Rows.Add("ИМЯ");

            _tblDetailsFirst.Columns.Add("НАИМЕНОВАНИЕ");
            _tblDetailsSecond.Columns.Add("НАИМЕНОВАНИЕ");


            for (int i = 0; i < _dwgObjcts.Count; i++)
            {
                int col = i + 1;
                _modified.Add(false);
                //string id = _dwgObjcts[i].ObjId.ToString();
                string sColNum = (i + 1).ToString();
                _tblNames.Columns.Add(sColNum);
                _tblDetailsFirst.Columns.Add(sColNum);
                _tblDetailsSecond.Columns.Add(sColNum);

                _tblNames.Rows[0][col] = _dwgObjcts[i].Number;
                _tblNames.Rows[1][col] = _dwgObjcts[i].Name;

                foreach (DataRow row in _dwgObjcts[i].Table_1.Rows)
                {
                    string name = row["item_name"].ToString();
                    string count = row["item_count"].ToString();
                    if (!detailsFirstTable.Contains(name))
                    {
                        detailsFirstTable.Add(name);
                        DataRow newrow = _tblDetailsFirst.NewRow();
                        _tblDetailsFirst.Rows.Add(newrow);
                    }
                    int rowFnd = detailsFirstTable.IndexOf(name);
                    _tblDetailsFirst.Rows[rowFnd][0] = name;
                    _tblDetailsFirst.Rows[rowFnd][col] = count;
                }

                foreach (DataRow row in _dwgObjcts[i].Table_2.Rows)
                {
                    string name = row["item_name"].ToString();
                    string count = row["item_count"].ToString();
                    if (!detailsSecondTable.Contains(name))
                    {
                        detailsSecondTable.Add(name);
                        DataRow newrow = _tblDetailsSecond.NewRow();
                        _tblDetailsSecond.Rows.Add(newrow);
                    }
                    int rowFnd = detailsSecondTable.IndexOf(name);
                    _tblDetailsSecond.Rows[rowFnd][0] = name;
                    _tblDetailsSecond.Rows[rowFnd][col] = count;
                }
            }
            IEnumerable<DataRow> query = (from row in _tblNames.AsEnumerable()
                                          select row).Union(from row2 in _tblDetailsFirst.AsEnumerable()
                                                            select row2).Union(from row3 in _tblDetailsSecond.AsEnumerable()
                                                                               select row3);
            _unionTable = query.CopyToDataTable<DataRow>();
        }

        private void SLGroupEditFrm_Load(object sender, EventArgs e)
        {
            ReloadTables();            
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                if (col.Index != 0)
                {
                    col.SortMode = DataGridViewColumnSortMode.NotSortable;
                    col.Width = 50;
                }
                else
                {
                    DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
                    columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;                    
                    columnHeaderStyle.BackColor = Color.LightGray;
                    col.DefaultCellStyle = columnHeaderStyle;
                    
                }
            }           

            dataGridView1.Columns[0].ReadOnly = true; 
            dataGridView1.Columns[0].Frozen = true; 
            dataGridView1.Rows[0].Frozen = true;

            //DataGridViewCellStyle myStyle = new DataGridViewCellStyle();
            //myStyle.BackColor = Color.LimeGreen;
            
            //dataGridView1.Rows[2].Cells[2].Style = myStyle;
        }

        private void ReloadTables()
        {
            bindingSource1.DataSource = _unionTable;
            dataGridView1.DataSource = bindingSource1;

            //dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            //dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int index = e.ColumnIndex - 1;
            _modified[index] = true;
            //m_modified = true;
        }
        

        private void ZoomToBlock(int index)
        {
            try
            {                
                ObjectId id = _dwgObjcts[index].ObjId;
                if (id != curRefid)
                {
                    ViewEntityPos(id);
                    curRefid = id;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public static void ViewEntityPos(ObjectId id)
        {
            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                //ed.UpdateTiledViewportsInDatabase();                
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                    Extents3d entityExtent = ent.GeometricExtents;

                    ViewportTable vpTbl = tr.GetObject(
                                                        db.ViewportTableId,
                                                        OpenMode.ForRead
                                                      ) as ViewportTable;

                    ViewportTableRecord viewportTableRec
                           = tr.GetObject(vpTbl["*Active"], OpenMode.ForWrite)
                                                       as ViewportTableRecord;

                    Matrix3d matWCS2DCS
                       = Matrix3d.PlaneToWorld(viewportTableRec.ViewDirection);

                    matWCS2DCS = Matrix3d.Displacement(
                        viewportTableRec.Target - Point3d.Origin) * matWCS2DCS;

                    matWCS2DCS = Matrix3d.Rotation
                                                (
                                                 -viewportTableRec.ViewTwist,
                                                 viewportTableRec.ViewDirection,
                                                 viewportTableRec.Target
                                                 ) * matWCS2DCS;

                    matWCS2DCS = matWCS2DCS.Inverse();

                    entityExtent.TransformBy(matWCS2DCS);

                    Point2d center = new Point2d(
                    (entityExtent.MaxPoint.X + entityExtent.MinPoint.X) * 0.5,
                    (entityExtent.MaxPoint.Y + entityExtent.MinPoint.Y) * 0.5);
                    //new
                    doc.Editor.Regen();
                    ent.Highlight();
                    Zoom(new Point3d(), new Point3d(), new Point3d(center.X, center.Y, 0), 1);

                    //viewportTableRec.CenterPoint = center;                    
                    tr.Commit();
                }
                //ed.UpdateTiledViewportsFromDatabase();


            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public static void Zoom(Point3d pMin, Point3d pMax, Point3d pCenter, double dFactor)
        {
            // Get the current document and database
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            int nCurVport = System.Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

            // Get the extents of the current space no points 
            // or only a center point is provided
            // Check to see if Model space is current
            if (acCurDb.TileMode == true)
            {
                if (pMin.Equals(new Point3d()) == true &&
                    pMax.Equals(new Point3d()) == true)
                {
                    pMin = acCurDb.Extmin;
                    pMax = acCurDb.Extmax;
                }
            }
            else
            {
                // Check to see if Paper space is current
                if (nCurVport == 1)
                {
                    // Get the extents of Paper space
                    if (pMin.Equals(new Point3d()) == true &&
                        pMax.Equals(new Point3d()) == true)
                    {
                        pMin = acCurDb.Pextmin;
                        pMax = acCurDb.Pextmax;
                    }
                }
                else
                {
                    // Get the extents of Model space
                    if (pMin.Equals(new Point3d()) == true &&
                        pMax.Equals(new Point3d()) == true)
                    {
                        pMin = acCurDb.Extmin;
                        pMax = acCurDb.Extmax;
                    }
                }
            }

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Get the current view
                using (ViewTableRecord acView = acDoc.Editor.GetCurrentView())
                {
                    Extents3d eExtents;

                    // Translate WCS coordinates to DCS
                    Matrix3d matWCS2DCS;
                    matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection);
                    matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) * matWCS2DCS;
                    matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist,
                                                   acView.ViewDirection,
                                                   acView.Target) * matWCS2DCS;

                    // If a center point is specified, define the min and max 
                    // point of the extents
                    // for Center and Scale modes
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        pMin = new Point3d(pCenter.X - (acView.Width / 2),
                                           pCenter.Y - (acView.Height / 2), 0);

                        pMax = new Point3d((acView.Width / 2) + pCenter.X,
                                           (acView.Height / 2) + pCenter.Y, 0);
                    }

                    // Create an extents object using a line
                    using (Line acLine = new Line(pMin, pMax))
                    {
                        eExtents = new Extents3d(acLine.Bounds.Value.MinPoint,
                                                 acLine.Bounds.Value.MaxPoint);
                    }

                    // Calculate the ratio between the width and height of the current view
                    double dViewRatio;
                    dViewRatio = (acView.Width / acView.Height);

                    // Tranform the extents of the view
                    matWCS2DCS = matWCS2DCS.Inverse();
                    eExtents.TransformBy(matWCS2DCS);

                    double dWidth;
                    double dHeight;
                    Point2d pNewCentPt;

                    // Check to see if a center point was provided (Center and Scale modes)
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        dWidth = acView.Width;
                        dHeight = acView.Height;

                        if (dFactor == 0)
                        {
                            pCenter = pCenter.TransformBy(matWCS2DCS);
                        }

                        pNewCentPt = new Point2d(pCenter.X, pCenter.Y);
                    }
                    else // Working in Window, Extents and Limits mode
                    {
                        // Calculate the new width and height of the current view
                        dWidth = eExtents.MaxPoint.X - eExtents.MinPoint.X;
                        dHeight = eExtents.MaxPoint.Y - eExtents.MinPoint.Y;

                        // Get the center of the view
                        pNewCentPt = new Point2d(((eExtents.MaxPoint.X + eExtents.MinPoint.X) * 0.5),
                                                 ((eExtents.MaxPoint.Y + eExtents.MinPoint.Y) * 0.5));
                    }

                    // Check to see if the new width fits in current window
                    if (dWidth > (dHeight * dViewRatio)) dHeight = dWidth / dViewRatio;

                    // Resize and scale the view
                    if (dFactor != 0)
                    {
                        acView.Height = dHeight * dFactor;
                        acView.Width = dWidth * dFactor;
                    }

                    // Set the center of the view
                    acView.CenterPoint = pNewCentPt;

                    // Set the current view
                    acDoc.Editor.SetCurrentView(acView);
                }

                // Commit the changes
                acTrans.Commit();
            }
        }

        private void SyncTables(System.Data.DataTable SourceTable, int colIndx, System.Data.DataTable ResultTable)
        {
            try
            {
                foreach (DataRow row in SourceTable.Rows)
                {
                    string name = row[0].ToString();
                    string count = row[colIndx].ToString();
                    double dcount;
                    if (count == "") dcount = 0;
                    else dcount = Convert.ToDouble(count, System.Globalization.CultureInfo.GetCultureInfo("en-US"));
                    //Три случая-деталь есть, изм. кол-во, детали не было, деталь убирается (кол-во становится равным нулю)!
                    int res = (from myRow in ResultTable.AsEnumerable()
                               where myRow.Field<string>("item_name") == name
                               select myRow).Count();
                    IEnumerable<DataRow> ExistRow = from myRow in ResultTable.AsEnumerable()
                                                where myRow.Field<string>("item_name") == name
                                                select myRow;
                    if (res == 0 && dcount>0)
                    {
                        //add new detail
                        DataRow newRow = ResultTable.NewRow();
                        newRow["item_name"] = name;
                        newRow["item_count"] = dcount;
                        ResultTable.Rows.Add(newRow);
                    }
                    if (res == 1)
                    {                        
                        //delete existing detail
                        if (dcount == 0)
                        {
                            ResultTable.Rows.Remove(ExistRow.First());
                        }
                        else
                        {
                            //set count
                            ExistRow.First()["item_count"] = dcount;
                        }                        
                    }
                    if (res < 0 || res > 1)
                    {
                        throw new IndexOutOfRangeException("Ошибка в количестве деталей " + name);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            try
            {
                for (int i = 0; i < _modified.Count; i++)
                {
                    if (_modified[i])
                    {
                        int colIndx = i + 1;
                        string Number = _unionTable.Rows[0][colIndx].ToString();
                        string Name = _unionTable.Rows[1][colIndx].ToString();
                        _dwgObjcts[i].Number = Number;
                        _dwgObjcts[i].Name = Name;
                        //Выделяем из единой таблицы таблицу2
                        int lBound = 1;
                        int UBound = _tblDetailsFirst.Rows.Count + lBound;
                        if (_tblDetailsFirst.Rows.Count > 0)
                        {
                            System.Data.DataTable tblDetailsFirst = _unionTable.AsEnumerable()
                                        .Where((row, nrow) => nrow > lBound && nrow <= UBound)
                                        .CopyToDataTable();
                            SyncTables(tblDetailsFirst, colIndx, _dwgObjcts[i].Table_1);
                        }
                        //Выделяем из единой таблицы таблицу3
                        lBound = UBound;
                        UBound = _tblDetailsSecond.Rows.Count + lBound;
                        if (_tblDetailsSecond.Rows.Count > 0)
                        {
                            System.Data.DataTable tblDetailsSecond = _unionTable.AsEnumerable()
                                        .Where((row, nrow) => nrow > lBound && nrow <= UBound)
                                        .CopyToDataTable();
                            SyncTables(tblDetailsSecond, colIndx, _dwgObjcts[i].Table_2);
                        }

                        _dwgObjcts[i].SaveXMLtoCADEntity(_dwgObjcts[i].ToXElement());
                        _dwgObjcts[i].SetAttributes();
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }
        }

        private void dataGridView2_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                DataGridView dgv = sender as DataGridView;
                if (dgv == null) return;
                dgv.Rows[e.RowIndex].ErrorText = "";
                double newDouble;
                if (dgv.Rows[e.RowIndex].IsNewRow) { return; }
                if (e.ColumnIndex == 0) { return; }

                string s_val = e.FormattedValue.ToString();
                if (s_val == "") { return; }

                if (!double.TryParse(e.FormattedValue.ToString(), NumberStyles.Float,
                    CultureInfo.CreateSpecificCulture("en-US"), out newDouble) || newDouble < 0)
                {
                    e.Cancel = true;
                    dgv.Rows[e.RowIndex].ErrorText = "Введите положительное число";
                }
            }
            catch (Exception ex)
            {                
               MessageBox.Show(ex.ToString());
            }
        }

        private void buttonZoom_Click(object sender, EventArgs e)
        {
            try
            {
                int col = this.dataGridView1.CurrentCell.ColumnIndex;
                if (col < 1) return;
                else ZoomToBlock(col - 1);
            }
            catch (Exception ex)
            {                
                MessageBox.Show(ex.ToString());
            }
        }

        private void dataGridView1_Paint(object sender, PaintEventArgs e)
        {            
            //Rectangle rec = (Rectangle)this.dataGridView1.DisplayRectangle;
            //int x0 = rec.Left;
            //int x1 = rec.Right;
            //int y = rec.Top + rec.Height / 2;
            //this.SuspendLayout();            
            //e.Graphics.DrawLine(new Pen(Color.Red, 2), new Point(x0, y), new Point(x1, y));           
            //this.ResumeLayout(false);
        }
 
    }
}
