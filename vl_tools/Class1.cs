using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System.Windows.Forms;
using System.Collections;
using System.Data;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using System.Globalization;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using LEP;
using System.Configuration;
using System.Data.SQLite;
using Autodesk.AutoCAD.Windows;

[assembly: CommandClass(typeof(vl_tools.Commands))]
namespace vl_tools
{
    public class ProfilePoint
    {
        public double Distance { get; set; }
        public double Height { get; set; }
        public string Decription { get; set; }
        public bool IsPipe { get; set; }
        public double Diameter { get; set; }
    }
    public class Abonent
    {
        public string TP { get; set; }
        public string NUMBER { get; set; }
        public string FIO { get; set; }
        public double POWER { get; set; }
        public string FIDER { get; set; }
        public string PROVOD { get; set; }
        public string LENGHT { get; set; }
        public Abonent()
        {
            TP = "";
            NUMBER = "";
            FIO = "";
            POWER = 0;
            FIDER = "";
            PROVOD = "";
            LENGHT = "";
        }

    }


    public class Commands : IExtensionApplication
    {

        #region IExtensionApplication Members
        static double PntAngle;
        private static string _TemplatePath;
        public static SQLiteConnection connection;
        //PicketViewerForm picketFrm;

        //панель с базой
        //protected static Autodesk.AutoCAD.Windows.PaletteSet vps = null;
        //public static UserControl1 usercntrl = null;

        public void Initialize()
        {
            DemandLoading.RegistryUpdate.RegisterForDemandLoading();
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage(this.GetType().Module.Name + " загружен и добавлен в автозагрузку\n");
            ed.WriteMessage("Для удаления из автозагрузки используйте команду vl_tools_REMAUTO\n");
            string appPath="";
            #region CONFIG
            try
            {
                // location update
                Assembly assem = Assembly.GetExecutingAssembly();
                string name = assem.ManifestModule.Name;
                appPath = assem.Location.Replace(name, "");
                _TemplatePath = appPath;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            #endregion

            #region DB
            try
            {
                // create connection
                string databaseName = appPath + "vl_data.db";
                connection =
                new SQLiteConnection(string.Format("Data Source={0};", databaseName));
                connection.Open();               
            }
            catch (System.Exception ex)
            {
                if (connection != null) connection.Close();
                MessageBox.Show(ex.ToString());
            }
            #endregion

            #region EVENTHANDLERS
            DocumentCollection dm = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            dm.DocumentCreated += new DocumentCollectionEventHandler(OnDocumentCreated);
            foreach (Document doc in dm)
            {
                doc.Editor.PointMonitor +=
                  new PointMonitorEventHandler(OnMonitorPoint);
                //SubscribeToImpliedSelectionChanged(doc);               
            }
            #endregion

            #region PALETTE
            //if (vps == null)
            //{
            //    vps = new Autodesk.AutoCAD.Windows.PaletteSet("Объемы",
            //    new Guid("{11AF4ED0-05CB-495C-96BB-8F4FE061E945}"));
            //    usercntrl = new UserControl1();
            //    usercntrl.Connection = connection;
            //    vps.Add("Объемы", usercntrl);

            //    vps.Size = new System.Drawing.Size(800, 600);
            //    vps.MinimumSize = new System.Drawing.Size(800, 600);
            //    vps.Dock = Autodesk.AutoCAD.Windows.DockSides.Left;
            //    vps.Location = new System.Drawing.Point(0, 0);
            //    vps.Style = PaletteSetStyles.ShowPropertiesMenu | PaletteSetStyles.ShowAutoHideButton | PaletteSetStyles.ShowCloseButton;
            //    vps.Visible = true;                                               
            //}
            #endregion

            //picketFrm = new PicketViewerForm();
            //Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(null, picketFrm, false);
        }

      
        public void Terminate()
        {
            try
            {
                //throw new NotImplementedException();
                if (connection != null) connection.Dispose();

                DocumentCollection dm = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                if (dm != null)
                {
                    Editor ed = dm.MdiActiveDocument.Editor;
                    ed.PointMonitor -= new PointMonitorEventHandler(OnMonitorPoint);                   
                    
                }
            }
            catch (System.Exception)
            {
                // The editor may no longer
                // be available on unload
            }
            
        }



        #endregion

        //public void SubscribeToImpliedSelectionChanged(Document doc)
        //{
        //    doc.ImpliedSelectionChanged += new EventHandler(this.acDoc_ImpliedSelectionChanged);
        //}

        //private void acDoc_ImpliedSelectionChanged(object sender, EventArgs e)
        //{
        //    usercntrl.ImpliedSelectionChanged(sender, e);            
        //}

        private void OnMonitorPoint(object sender, PointMonitorEventArgs e)
        {
            int round = 2;
            int offset_round = 2;
            try
            {
                if (!e.Context.PointComputed) return;
                Point3d pnt = e.Context.ComputedPoint;

                // Get the current document and database
                Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                //check current space                
                using (Transaction tr = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTableRecord space = (BlockTableRecord)tr.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForRead);
                    if (space.Name.ToUpper() != BlockTableRecord.ModelSpace || space.IsDynamicBlock)
                    {
                        //ed.WriteMessage("Программа может быть выполнена только в простанстве модели!\n");
                        tr.Commit();
                        return;
                    }
                    tr.Commit();
                }

                VLFileOptions FileOpt = new VLFileOptions();
                if (!FileOpt.TryToRead()) return;
                if (!FileOpt.VL_PROFILE_MONITOR) return;//monitor Profile OFF
                VLPicketClass vlPicketObj = new VLPicketClass(FileOpt);
                vlPicketObj.Calc(pnt, round, offset_round);
                //picketFrm.SetText(vlPicketObj._sPicket, vlPicketObj._sKilometer, vlPicketObj._sSide + " " + vlPicketObj._sOffset);
                string statusstr = string.Format("{0,-20} {1,-20} {2,-10} {3,-10}", vlPicketObj._sPicket, vlPicketObj._sKilometer, vlPicketObj._sSide + ":", vlPicketObj._sOffset);
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("MODEMACRO", statusstr);
                //e.AppendToolTipText(vlPicketObj.ToString());
            }
            catch (System.Exception)
            {
                //System.Diagnostics.Debug.Print(ex.StackTrace);
                //MessageBox.Show(ex.ToString());
            }
        }

        private void OnDocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            e.Document.Editor.PointMonitor += new PointMonitorEventHandler(OnMonitorPoint);
            //SubscribeToImpliedSelectionChanged(e.Document);
        }

        //private Polyline GetCurrentTrace()
        //{
        //    try
        //    {
        //        Database db = HostApplicationServices.WorkingDatabase;
        //        using (Transaction trans =db.TransactionManager.StartTransaction())
        //        {
        //            // Find the NOD in the database
        //            DBDictionary nod = (DBDictionary)trans.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
        //            if (!nod.Contains(nodName)) return null;
        //            ObjectId myDataId = nod.GetAt(nodName);
        //            Xrecord readBack = (Xrecord)trans.GetObject(myDataId, OpenMode.ForRead);
        //            foreach (TypedValue value in readBack.Data)
        //            {
        //                System.Diagnostics.Debug.Print("===== OUR DATA: " + value.TypeCode.ToString() + ". " + value.Value.ToString());
        //            }
        //            trans.Commit();
        //        } // using
        //        return null;
        //    }
        //    catch (System.Exception)
        //    {
        //        return null;
        //    }
        //}

        //[CommandMethod("vl_SET_TRACE")]
        //public void cmdSetTrace()
        //{
        //    Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
        //    Database db = HostApplicationServices.WorkingDatabase;
        //    TypedValue[] values = new TypedValue[] { new TypedValue(0, "LWPOLYLINE") };
        //    SelectionFilter filter = new SelectionFilter(values);
        //    PromptSelectionOptions opts = new PromptSelectionOptions();

        //    //opts.MessageForRemoval = "\nMust be a type of Block!";
        //    opts.MessageForAdding = "\nSelect a polyline: ";
        //    //opts.PrepareOptionalDetails = true;
        //    opts.SingleOnly = true;
        //    //opts.SinglePickInSpace = true;
        //    opts.AllowDuplicates = false;
        //    PromptSelectionResult res = default(PromptSelectionResult);
        //    res = ed.GetSelection(opts, filter);
        //    if (res.Status != PromptStatus.OK) return;
        //    try
        //    {
        //        using (Transaction tr = db.TransactionManager.StartTransaction())
        //        {
        //            BlockTable acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        //            BlockTableRecord acBlkTblRec = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
        //            Polyline pline = tr.GetObject(res.Value[0].ObjectId, OpenMode.ForRead, false) as Polyline;
        //            if (pline == null) throw new NullReferenceException();
        //            tr.Commit();
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
        //    }

        //}

        [CommandMethod("vl_tools_REMAUTO")]
        public void cmdRemoveAuto()
        {
            DemandLoading.RegistryUpdate.UnregisterForDemandLoading();
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage(this.GetType().Module.Name + " удален из автозагрузки\n");
        }

        [CommandMethod("sum", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void cmdCalc()
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;
            PromptSelectionOptions opt = new PromptSelectionOptions();
            PromptSelectionResult res;
            res = ed.GetSelection(opt);
            if (res.Status != PromptStatus.OK) return;
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {                    
                    BlockTableRecord acBlkTblRec = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                    List<double> values = new List<double>();
                    //double sum = 0;
                    string s = "";
                    foreach (SelectedObject selPl in res.Value)
                    {
                        DBObject obj = tr.GetObject(selPl.ObjectId, OpenMode.ForRead, false);
                        if (obj.GetType() == typeof(Autodesk.AutoCAD.DatabaseServices.MText))
                        {
                            MText mtxt = obj as MText;
                            s = mtxt.Text;
                            double val = GetDoubleFString(s);
                            values.Add(val);
                            //sum += Convert.ToDouble(s);
                        }
                        if (obj.GetType() == typeof(Autodesk.AutoCAD.DatabaseServices.DBText))
                        {
                            DBText txt = obj as DBText;
                            s = txt.TextString;
                            double val = GetDoubleFString(s);
                            values.Add(val);
                            //sum += Convert.ToDouble(s);
                        }
                        if (obj is Autodesk.AutoCAD.DatabaseServices.Dimension)
                        {
                            Dimension dim = obj as Dimension;
                            if (dim.DimensionText == "")
                            {
                                //sum += dim.Measurement;
                                int dimdec = dim.Dimdec;
                                double val = Math.Round(dim.Measurement, dimdec);
                                values.Add(val);
                            }
                            else
                            {
                                double val = GetDoubleFString(dim.DimensionText);
                                //sum += Convert.ToDouble(s);
                                values.Add(val);
                            }
                        }

                    }
                    //create excel
                    Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                    //Книга.
                    ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
                    //Таблица.
                    ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
                    int row = 1;
                    foreach (double val in values)
                    {
                        ObjWorkSheet.Cells[row, 1] = val;
                        row++;
                    }
                    ObjWorkSheet.Cells[row, 1] = "=СУММ(A1:A" + (row - 1).ToString() + ")";
                    ObjExcel.Visible = true;
                    ObjExcel.UserControl = true;

                    //Собираем мусор
                    ObjExcel = null;
                    ObjWorkBook = null;
                    ObjWorkSheet = null;
                    GC.Collect();

                    tr.Commit();
                    //ed.WriteMessage(sum.ToString());
                }
            }

            catch (System.Exception ex)
            {
                ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
            }
        }

        public static double GetDoubleFString(string s)
        {
            s = s.Replace(',', '.');

            string result = "";
            bool start = false;

            for (int i = 0; i < s.Length; i++)
            {
                if ((s[i] > 47 && s[i] < 58) || (s[i] == 46) || (s[i] == 45))
                {
                    if (!start) start = true;
                    result += s[i];
                }
                else
                {
                    if (start) break;
                }
            }
            try
            {
                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = ".";
                return Convert.ToDouble(result, provider);
            }
            catch (System.Exception)
            {
                throw;
            }
        }

        //[CommandMethod("provis")]
        //public void cmdProvis()
        //{
        //    Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
        //    Database db = HostApplicationServices.WorkingDatabase;
        //    try
        //    {
        //        using (Transaction tr = db.TransactionManager.StartTransaction())
        //        {
        //            BlockTableRecord acBlkTblRec = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);                    
        //        }
        //    }

        //    catch (System.Exception ex)
        //    {
        //        ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
        //    }
        //}

        //private double GetDoubleFString(string s)
        //{
        //    s = s.Replace(',', '.');

        //    string result = "";
        //    bool start = false;

        //    for (int i = 0; i < s.Length; i++)
        //    {
        //        if ((s[i] > 47 && s[i] < 58) || (s[i] == 46))
        //        {
        //            if (!start) start = true;
        //            result += s[i];
        //        }
        //        else
        //        {
        //            if (start) break;
        //        }
        //    }
        //    return Convert.ToDouble(result);
        //}

        [CommandMethod("pldim")]
        public void cmdPLineDimSpeed()
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;
            TypedValue[] values = new TypedValue[] { new TypedValue(0, "LWPOLYLINE") };
            SelectionFilter filter = new SelectionFilter(values);
            PromptSelectionOptions opts = new PromptSelectionOptions();

            //opts.MessageForRemoval = "\nMust be a type of Block!";
            opts.MessageForAdding = "\nSelect a polyline: ";
            //opts.PrepareOptionalDetails = true;
            //opts.SingleOnly = true;
            //opts.SinglePickInSpace = true;
            //opts.AllowDuplicates = false;
            PromptSelectionResult res = default(PromptSelectionResult);
            res = ed.GetSelection(opts, filter);
            if (res.Status != PromptStatus.OK)
                return;
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    foreach (SelectedObject selPl in res.Value)
                    {
                        Polyline pline = tr.GetObject(selPl.ObjectId, OpenMode.ForRead, false) as Polyline;
                        for (int i = 1; i < pline.NumberOfVertices; i++)
                        {
                            Point2d pnt1 = pline.GetPoint2dAt(i - 1);
                            Point2d pnt2 = pline.GetPoint2dAt(i);
                            AlignedDimension drdim = new AlignedDimension(new Point3d(pnt1.X, pnt1.Y, 0),
                                new Point3d(pnt2.X, pnt2.Y, 0), new Point3d(pnt1.X, pnt1.Y, 0), "", ObjectId.Null);
                            acBlkTblRec.AppendEntity(drdim);
                            tr.AddNewlyCreatedDBObject(drdim, true);
                        }

                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
            }

        }

        [CommandMethod("mins")]
        public void cmdMultiInsert()
        {
            try
            {
                var w = Clipboard.GetData(DataFormats.Text);//Читаем
                string str = w.ToString();
                str = str.Replace("\n", "");
                string[] lines;
                lines = str.Split('\r');
                List<string> array = new List<string>();
                foreach (string s in lines)
                {
                    string[] cols;
                    cols = s.Split('\t');
                    array.AddRange(cols);
                }
                array.Remove("");

                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Database db = HostApplicationServices.WorkingDatabase;

                foreach (string s in array)
                {
                    PromptEntityOptions opt = new PromptEntityOptions("Select text, mtext or dimension");
                    opt.SetRejectMessage("Selected object not a text, mtext or dimension");
                    opt.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.MText), false);
                    //opt.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.AlignedDimension), false);
                    opt.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.DBText), false);
                    opt.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Dimension), false);

                    PromptEntityResult res = ed.GetEntity(opt);
                    if (res.Status != PromptStatus.OK) return;
                    try
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockTable acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                            BlockTableRecord acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                            DBObject obj = tr.GetObject(res.ObjectId, OpenMode.ForWrite, false);
                            if (obj.GetType() == typeof(Autodesk.AutoCAD.DatabaseServices.MText))
                            {
                                MText mtxt = obj as MText;
                                mtxt.Contents = s;
                            }
                            if (obj.GetType() == typeof(Autodesk.AutoCAD.DatabaseServices.DBText))
                            {
                                DBText txt = obj as DBText;
                                txt.TextString = s;
                            }
                            if (obj is Autodesk.AutoCAD.DatabaseServices.Dimension)
                            {
                                Dimension dim = obj as Dimension;
                                dim.DimensionText = s;
                            }

                            tr.Commit();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
            }
        }

        [CommandMethod("tblins")]
        public void cmdTableInsert()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            try
            {
                var w = Clipboard.GetData(DataFormats.Text);//Читаем
                string str = w.ToString();
                str = str.Replace("\n", "");
                string[] lines;
                lines = str.Split('\r');
                int nTblRows = lines.Length;

                //List<string> array = new List<string>();
                //foreach (string s in lines)
                //{
                //    string[] cols;
                //    cols = s.Split('\t');
                //    array.AddRange(cols);
                //}
                //array.Remove("");

                var opt = new PromptEntityOptions("\nSelect table to update");
                opt.SetRejectMessage("\nEntity is not a table.");
                opt.AddAllowedClass(typeof(Table), false);

                var per = ed.GetEntity(opt);
                if (per.Status != PromptStatus.OK) return;

                using (var tr = db.TransactionManager.StartTransaction())
                {

                    var obj = tr.GetObject(per.ObjectId, OpenMode.ForWrite);
                    var tb = obj as Table;

                    if (tb != null)
                    {
                        // The table must be open for write
                        tb.UpgradeOpen();

                        int nrows = tb.Rows.Count;
                        int ncols = tb.Columns.Count;

                        PromptIntegerOptions intopt = new PromptIntegerOptions("\nВведите начальный столбец: ");
                        intopt.LowerLimit = 1;
                        intopt.UpperLimit = ncols;
                        intopt.DefaultValue = 1;
                        PromptIntegerResult res = ed.GetInteger(intopt);
                        if (!(res.Status == PromptStatus.OK)) { ed.WriteMessage("Programm was cancelled"); return; }
                        int init_col = res.Value;

                        intopt.Message = "\nВведите начальную строку: ";
                        intopt.UpperLimit = nrows;
                        res = ed.GetInteger(intopt);
                        if (!(res.Status == PromptStatus.OK)) { ed.WriteMessage("Programm was cancelled"); return; }
                        int init_row = res.Value;

                        int col = init_col;
                        int row = init_row;

                        foreach (string s in lines)
                        {
                            string[] cols;
                            cols = s.Split('\t');
                            foreach (string value in cols)
                            {
                                tb.Cells[row - 1, col - 1].TextString = value;
                                col++;
                                if (col > ncols) break;
                            }
                            row++;
                            if (row > nrows) break;
                            col = init_col;
                        }
                        //foreach (string value in array)
                        //{
                        //    tb.Cells[row-1, col-1].TextString = value;
                        //    col++;
                        //    if (col > ncols)
                        //    {
                        //        row++;
                        //        col = init_col;
                        //        if (row > nrows) break;
                        //    }
                        //}
                        //for (int i = row; i < nrows; i++)
                        //{
                        //    for (int j = col; j < ncols; j++)
                        //    {

                        //    }
                        //}
                    }
                    tr.Commit();

                }
            }//end try
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("\nException: {0}", ex.Message);
            }
            catch (System.NullReferenceException ex)
            {
                ed.WriteMessage("\nException: {0}", ex.Message);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nException: {0}", ex.Message);
            }
        }


        [CommandMethod("BlockAttsUpdate")]
        public void cmdBlockAttributesUpdate()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            try
            {
                if (!Clipboard.ContainsText())
                {
                    ed.WriteMessage("Буфер обмена не содержит текст. Команда будет завершена.\n");
                    return;
                }
                var w = Clipboard.GetData(DataFormats.Text);//Читаем
                string str = w.ToString();
                str = str.Replace("\n", "");
                string[] lines;
                lines = str.Split('\r');
                //remove last split element ""
                Array.Resize(ref lines, lines.Length - 1);
                List<string> array = new List<string>();

                string[] fields = lines[0].Split('\t');

                foreach (string s in lines)
                {
                    string[] cols;
                    cols = s.Split('\t');

                    array.AddRange(cols);
                }

                //array.Remove("");
                int ncols = fields.Length;
                array.RemoveRange(0, ncols);
                int nrows = array.Count / ncols;

                //Select Blocks
                #region SELECTING
                //Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                //Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                PromptSelectionOptions selOpt = new PromptSelectionOptions();
                selOpt.MessageForAdding = "Выберите блоки абонентов: ";
                TypedValue[] values = { new TypedValue((int)DxfCode.Start, "INSERT"), /*new TypedValue((int)DxfCode.BlockName, "DETAIL_W")*/ };
                SelectionFilter sfilter = new SelectionFilter(values);
                selOpt.AllowDuplicates = false;
                PromptSelectionResult sset = ed.GetSelection(selOpt, sfilter);
                if (!(sset.Status == PromptStatus.OK)) { ed.WriteMessage("Программа отменена пользователем.\n"); return; }

                ObjectId[] objIds = sset.Value.GetObjectIds();
                Database dbCurrent = HostApplicationServices.WorkingDatabase;
                #endregion
                #region FILL_TABLE


                List<Abonent> abonents = new List<Abonent>();
                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in objIds)
                    {
                        BlockReference bRef = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                        AttributeCollection attcol = bRef.AttributeCollection;                        
                        Abonent ab = new Abonent();
                        foreach (ObjectId att in attcol)
                        {
                            AttributeReference atRef = (AttributeReference)tr.GetObject(att, OpenMode.ForRead);

                            //if (atRef.Tag == "ТП")
                            //{
                            //    if (atRef.TextString == "")
                            //    {
                            //        string errstr = "В блоке абонента " + " " + ab.NUMBER + " " + ab.FIO + " x=" + bRef.Position.X + " y=" + bRef.Position.Y + " не указана ТП!";
                            //        MessageBox.Show(errstr, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            //        return;
                            //    }
                            //    ab.TP = atRef.TextString;
                            //    fnd = true;
                            //}
                            if (atRef.Tag == "НОМЕР_ЗАЯВКИ")
                            {
                                //ab.NUMBER = atRef.TextString;
                                foreach (ObjectId att2 in attcol)
                                {
                                    AttributeReference atRef2 = (AttributeReference)tr.GetObject(att2, OpenMode.ForWrite);
                                    if (atRef2.Tag == "МОЩНОСТЬ")
                                    {
                                        for (int i = 0; i < nrows; i++)
                                        {
                                            Hashtable tbl = new Hashtable();

                                            for (int j = 0; j < ncols; j++)
                                            {
                                                string fld = fields[j];
                                                string value = array[i * ncols + j];
                                                tbl.Add(fld, value);
                                            }
                                            if (tbl["НОМЕР_ЗАЯВКИ"].ToString()== atRef.TextString)
                                            {
                                                atRef2.TextString = tbl["МОЩНОСТЬ"].ToString();
                                                ed.WriteMessage("Для абонента " + atRef.TextString + " установлена мощность " + atRef2.TextString + "\n");
                                            }                                         
                                        }

                                    }

                                    if (atRef2.Tag == "ПРИМЕЧАНИЕ")
                                    {
                                        for (int i = 0; i < nrows; i++)
                                        {
                                            Hashtable tbl = new Hashtable();

                                            for (int j = 0; j < ncols; j++)
                                            {
                                                string fld = fields[j];
                                                string value = array[i * ncols + j];
                                                tbl.Add(fld, value);
                                            }
                                            if (tbl["НОМЕР_ЗАЯВКИ"].ToString() == atRef.TextString)
                                            {
                                                atRef2.TextString = tbl["ПРИМЕЧАНИЕ"].ToString();
                                                ed.WriteMessage("Для абонента " + atRef.TextString + " установлено примечание " + atRef2.TextString + "\n");
                                            }
                                        }

                                    }

                                    if (atRef2.Tag == "ОФЗ")
                                    {
                                        for (int i = 0; i < nrows; i++)
                                        {
                                            Hashtable tbl = new Hashtable();

                                            for (int j = 0; j < ncols; j++)
                                            {
                                                string fld = fields[j];
                                                string value = array[i * ncols + j];
                                                tbl.Add(fld, value);
                                            }
                                            if (tbl["НОМЕР_ЗАЯВКИ"].ToString() == atRef.TextString)
                                            {
                                                atRef2.TextString = tbl["ОФЗ"].ToString();
                                                ed.WriteMessage("Для абонента " + atRef.TextString + " установлен аттрибут ОФЗ " + atRef2.TextString + "\n");
                                            }
                                        }

                                    }

                                    if (atRef2.Tag == "ФИО")
                                    {
                                        for (int i = 0; i < nrows; i++)
                                        {
                                            Hashtable tbl = new Hashtable();

                                            for (int j = 0; j < ncols; j++)
                                            {
                                                string fld = fields[j];
                                                string value = array[i * ncols + j];
                                                tbl.Add(fld, value);
                                            }
                                            if (tbl["НОМЕР_ЗАЯВКИ"].ToString() == atRef.TextString)
                                            {
                                                atRef2.TextString = tbl["ФИО"].ToString();
                                                ed.WriteMessage("Для абонента " + atRef.TextString + " установлен ФИО " + atRef2.TextString + "\n");
                                            }
                                        }

                                    }
                                }
                                //if (atRef.TextString == "")
                                //{
                                //    string errstr = "В блоке абонента " + " " + ab.NUMBER + " " + ab.FIO + " x=" + bRef.Position.X + " y=" + bRef.Position.Y + " не указан номер заявки!";
                                //    MessageBox.Show(errstr, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                //    return;
                                //}
                            }
                            //if (atRef.Tag == "ФИО")
                            //{
                            //    ab.FIO = atRef.TextString;
                            //    if (atRef.TextString == "")
                            //    {
                            //        string errstr = "В блоке абонента " + " " + ab.NUMBER + " " + ab.FIO + " x=" + bRef.Position.X + " y=" + bRef.Position.Y + " не указана фамилия!";
                            //        MessageBox.Show(errstr, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            //        return;
                            //    }
                            //}
                        }
                        
                    }
                    tr.Commit();
                }
                #endregion FILL_TABLE 
                

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //[CommandMethod("VL_SHOW_PALETTESET")]
        //public void cmdVLShowPaletteSet()
        //{
        //    if (vps != null)
        //    {
        //        vps.Visible = true;
        //    }
        //}


        [CommandMethod("BlAtsins")]
        public void cmdBlockAttributesInsert()
        {
            //Вставляет блоки с аттрибутами из буфера обмена
            //Имена полей скопированного текста и аттрибутов должны совпадать
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            try
            {
                if (!Clipboard.ContainsText())
                {
                    ed.WriteMessage("Буфер обмена не содержит текст. Команда будет завершена.\n");
                    return;
                }

                var w = Clipboard.GetData(DataFormats.Text);//Читаем
                string str = w.ToString();
                str = str.Replace("\n", "");
                string[] lines;
                lines = str.Split('\r');
                //remove last split element ""
                Array.Resize(ref lines, lines.Length - 1);
                List<string> array = new List<string>();
                
                string[] fields = lines[0].Split('\t');
                
                foreach (string s in lines)
                {
                    string[] cols;
                    cols = s.Split('\t');
                    
                    array.AddRange(cols);
                }

                //array.Remove("");
                int ncols = fields.Length;
                array.RemoveRange(0,ncols);
                int nrows = array.Count / ncols;

                //Запрос типа блока
                //PromptStringOptions opt = new PromptStringOptions("Введите имя блока: ");
                //opt.AllowSpaces = true;
                //PromptResult bres = ed.GetString(opt);
                //if (bres.Status != PromptStatus.OK) return;
                //string blockName = bres.StringResult;
                PromptEntityOptions opt = new PromptEntityOptions("Выберите блок: ");
                opt.SetRejectMessage("Выбран не блок!");
                opt.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.BlockReference), false);
                 PromptEntityResult res = ed.GetEntity(opt);
                if (res.Status != PromptStatus.OK) return;                
                ObjectId blkId = res.ObjectId;

                //Get blockname
                string blockName = "";
                BlockTableRecord btr = null;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockReference blkRef = (BlockReference)tr.GetObject(blkId, OpenMode.ForRead);
                    btr = (BlockTableRecord)tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead);
                    blockName = btr.Name;
                    btr.Dispose();
                }

                //Get method
                string[] skwrds = new string[] { "1", "2"};
                PromptResult KwRes = ed.GetKeywords("Метод построения: 1-вручную, 2-по координатам", skwrds);
                if (KwRes.Status != PromptStatus.OK) return;
                
                if(KwRes.StringResult=="1")
                {
                    for (int i = 0; i < nrows; i++)
                    {
                        Hashtable tbl = new Hashtable();

                        for (int j = 0; j < ncols; j++)
                        {
                            string fld = fields[j];
                            string value = array[i * ncols + j];
                            tbl.Add(fld, value);
                        }

                        BlockReference br = this.BlockJig(blockName, tbl);
                        if (br == null)
                        {
                            ed.WriteMessage("Прервано пользователем\n");
                            return;
                        }
                    }
                }

                if (KwRes.StringResult == "2")
                {
                    for (int i = 0; i < nrows; i++)
                    {
                        Hashtable tbl = new Hashtable();
                        double Xcoord = 0;
                        double Ycoord = 0;
                        for (int j = 0; j < ncols; j++)
                        {
                            string fld = fields[j];   
                            string value = array[i * ncols + j];
                            if (fld == "X")
                            {
                                Xcoord = GetDoubleFString(value);
                            }
                            if (fld == "Y")
                            {
                                Ycoord = GetDoubleFString(value);
                            }
                            tbl.Add(fld, value);
                        }
                        Point3d pnt1 = new Point3d(Xcoord, Ycoord, 0);
                        this.BlockWithAttributesInsert(blockName, tbl, pnt1);
                    }
                }




                //Point2d pnt1 = new Point2d(0, 0);
                //BlockReference br = new BlockReference(new Point3d(pnt1.X, pnt1.Y-i*20, 0), btr.ObjectId);
                //space.AppendEntity(br);
                //Dictionary<ObjectId, AttInfo> attInfo = new Dictionary<ObjectId, AttInfo>();
                //if (btr.HasAttributeDefinitions)
                //{
                //    foreach (ObjectId id in btr)
                //    {
                //        DBObject obj =
                //            tr.GetObject(id, OpenMode.ForRead);
                //        AttributeDefinition ad =
                //            obj as AttributeDefinition;

                //        if (ad != null && !ad.Constant)
                //        {
                //            AttributeReference ar =
                //                new AttributeReference();

                //            ar.SetAttributeFromBlock(ad, br.BlockTransform);
                //            ar.Position =
                //                ad.Position.TransformBy(br.BlockTransform);

                //            if (ad.Justify != AttachmentPoint.BaseLeft)
                //            {
                //                ar.AlignmentPoint =
                //                    ad.AlignmentPoint.TransformBy(br.BlockTransform);
                //            }
                //            if (ar.IsMTextAttribute)
                //            {
                //                ar.UpdateMTextAttribute();
                //            }

                //            ar.TextString = ad.TextString;

                //            ObjectId arId =
                //                br.AttributeCollection.AppendAttribute(ar);
                //            tr.AddNewlyCreatedDBObject(ar, true);

                //            // Initialize our dictionary with the ObjectId of
                //            // the attribute reference + attribute definition info

                //            attInfo.Add(
                //                arId,
                //                new AttInfo(
                //                ad.Position,
                //                ad.AlignmentPoint,
                //                ad.Justify != AttachmentPoint.BaseLeft
                //                )
                //            );
                //        }
                //    }
                //}

                //tr.AddNewlyCreatedDBObject(br, true);






                //foreach (SelectedObject selPl in res.Value)
                //{
                //    Polyline pline = tr.GetObject(selPl.ObjectId, OpenMode.ForRead, false) as Polyline;
                //    for (int i = 0; i < pline.NumberOfVertices; i++)
                //    {
                //        Point2d pnt1 = pline.GetPoint2dAt(i);

                //        BlockReference br = new BlockReference(new Point3d(pnt1.X, pnt1.Y, 0), btr.ObjectId);
                //        space.AppendEntity(br);
                //        tr.AddNewlyCreatedDBObject(br, true);

                //        Dictionary<ObjectId, AttInfo> attInfo = new Dictionary<ObjectId, AttInfo>();
                //        if (btr.HasAttributeDefinitions)
                //        {
                //            foreach (ObjectId id in btr)
                //            {
                //                DBObject obj =
                //                    tr.GetObject(id, OpenMode.ForRead);
                //                AttributeDefinition ad =
                //                    obj as AttributeDefinition;

                //                if (ad != null && !ad.Constant)
                //                {
                //                    AttributeReference ar =
                //                        new AttributeReference();

                //                    ar.SetAttributeFromBlock(ad, br.BlockTransform);
                //                    ar.Position =
                //                        ad.Position.TransformBy(br.BlockTransform);

                //                    if (ad.Justify != AttachmentPoint.BaseLeft)
                //                    {
                //                        ar.AlignmentPoint =
                //                            ad.AlignmentPoint.TransformBy(br.BlockTransform);
                //                    }
                //                    if (ar.IsMTextAttribute)
                //                    {
                //                        ar.UpdateMTextAttribute();
                //                    }

                //                    ar.TextString = ad.TextString;

                //                    ObjectId arId =
                //                        br.AttributeCollection.AppendAttribute(ar);
                //                    tr.AddNewlyCreatedDBObject(ar, true);

                //                    // Initialize our dictionary with the ObjectId of
                //                    // the attribute reference + attribute definition info

                //                    attInfo.Add(
                //                        arId,
                //                        new AttInfo(
                //                        ad.Position,
                //                        ad.AlignmentPoint,
                //                        ad.Justify != AttachmentPoint.BaseLeft
                //                        )
                //                    );
                //                }
                //            }
                //        }
                //    }

                //}



            }//end try
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("\nException: {0}", ex.Message);
            }
            catch (System.NullReferenceException ex)
            {
                ed.WriteMessage("\nException: {0}", ex.Message);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nException: {0}", ex.Message);
            }
        }
        private BlockReference BlockWithAttributesInsert(string BlockName, Hashtable HAttributes, Point3d insPnt)
        {
            Autodesk.AutoCAD.ApplicationServices.Document doc =
              Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            BlockReference br = null;

            Transaction tr =
              doc.TransactionManager.StartTransaction();
            using (tr)
            {
                BlockTable bt =
                  (BlockTable)tr.GetObject(
                    db.BlockTableId,
                    OpenMode.ForRead
                  );

                if (!bt.Has(BlockName))
                {
                    throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.InvalidBlockName,
                      "\nBlock \"" + BlockName + "\" not found.");
                }

                BlockTableRecord space =
                  (BlockTableRecord)tr.GetObject(
                    db.CurrentSpaceId,
                    OpenMode.ForWrite
                  );

                BlockTableRecord btr =
                  (BlockTableRecord)tr.GetObject(
                    bt[BlockName],
                    OpenMode.ForRead);

                // Block needs to be inserted to current space before
                // being able to append attribute to it

                br = new BlockReference(insPnt, btr.ObjectId);
                space.AppendEntity(br);
                tr.AddNewlyCreatedDBObject(br, true);

                Dictionary<ObjectId, AttInfo> attInfo =
                  new Dictionary<ObjectId, AttInfo>();

                if (btr.HasAttributeDefinitions)
                {
                    foreach (ObjectId id in btr)
                    {
                        DBObject obj =
                          tr.GetObject(id, OpenMode.ForRead);
                        AttributeDefinition ad =
                          obj as AttributeDefinition;

                        if (ad != null && !ad.Constant)
                        {
                            AttributeReference ar =
                              new AttributeReference();

                            ar.SetAttributeFromBlock(ad, br.BlockTransform);
                            ar.Position =
                              ad.Position.TransformBy(br.BlockTransform);

                            if (ad.Justify != AttachmentPoint.BaseLeft)
                            {
                                ar.AlignmentPoint =
                                  ad.AlignmentPoint.TransformBy(br.BlockTransform);
                            }
                            if (ar.IsMTextAttribute)
                            {
                                ar.UpdateMTextAttribute();
                            }

                            if (HAttributes != null && HAttributes.Contains(ad.Tag))
                            {
                                ar.TextString = HAttributes[ad.Tag].ToString();

                            }
                            else
                            {
                                ar.TextString = ad.TextString;
                            }

                            ObjectId arId =
                              br.AttributeCollection.AppendAttribute(ar);
                            tr.AddNewlyCreatedDBObject(ar, true);

                            // Initialize our dictionary with the ObjectId of
                            // the attribute reference + attribute definition info

                            attInfo.Add(
                              arId,
                              new AttInfo(
                                ad.Position,
                                ad.AlignmentPoint,
                                ad.Justify != AttachmentPoint.BaseLeft
                              )
                            );
                        }
                    }
                } 
                // Commit changes if user accepted, otherwise discard
                tr.Commit();

            }
            return br;
        }
        private BlockReference BlockJig(string BlockName, Hashtable HAttributes)
        {
            Autodesk.AutoCAD.ApplicationServices.Document doc =
              Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            BlockReference br = null;

            Transaction tr =
              doc.TransactionManager.StartTransaction();
            using (tr)
            {
                BlockTable bt =
                  (BlockTable)tr.GetObject(
                    db.BlockTableId,
                    OpenMode.ForRead
                  );

                if (!bt.Has(BlockName))
                {
                    throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.InvalidBlockName,
                      "\nBlock \"" + BlockName + "\" not found.");
                }

                BlockTableRecord space =
                  (BlockTableRecord)tr.GetObject(
                    db.CurrentSpaceId,
                    OpenMode.ForWrite
                  );

                BlockTableRecord btr =
                  (BlockTableRecord)tr.GetObject(
                    bt[BlockName],
                    OpenMode.ForRead);

                // Block needs to be inserted to current space before
                // being able to append attribute to it

                br = new BlockReference(new Point3d(), btr.ObjectId);
                space.AppendEntity(br);
                tr.AddNewlyCreatedDBObject(br, true);

                Dictionary<ObjectId, AttInfo> attInfo =
                  new Dictionary<ObjectId, AttInfo>();

                if (btr.HasAttributeDefinitions)
                {
                    foreach (ObjectId id in btr)
                    {
                        DBObject obj =
                          tr.GetObject(id, OpenMode.ForRead);
                        AttributeDefinition ad =
                          obj as AttributeDefinition;

                        if (ad != null && !ad.Constant)
                        {
                            AttributeReference ar =
                              new AttributeReference();

                            ar.SetAttributeFromBlock(ad, br.BlockTransform);
                            ar.Position =
                              ad.Position.TransformBy(br.BlockTransform);

                            if (ad.Justify != AttachmentPoint.BaseLeft)
                            {
                                ar.AlignmentPoint =
                                  ad.AlignmentPoint.TransformBy(br.BlockTransform);
                            }
                            if (ar.IsMTextAttribute)
                            {
                                ar.UpdateMTextAttribute();
                            }

                            if (HAttributes != null && HAttributes.Contains(ad.Tag))
                            {
                                ar.TextString = HAttributes[ad.Tag].ToString();

                            }
                            else
                            {
                                ar.TextString = ad.TextString;
                            }

                            ObjectId arId =
                              br.AttributeCollection.AppendAttribute(ar);
                            tr.AddNewlyCreatedDBObject(ar, true);

                            // Initialize our dictionary with the ObjectId of
                            // the attribute reference + attribute definition info

                            attInfo.Add(
                              arId,
                              new AttInfo(
                                ad.Position,
                                ad.AlignmentPoint,
                                ad.Justify != AttachmentPoint.BaseLeft
                              )
                            );
                        }
                    }
                }
                // Run the jig

                BlockJig myJig = new BlockJig(tr, br, attInfo);

                if (myJig.Run() != PromptStatus.OK)
                    return null;

                // Commit changes if user accepted, otherwise discard
                tr.Commit();

            }
            return br;
        }
        /// <summary>
        /// Inserts all attributreferences
        /// </summary>
        /// <param name="blkRef">Blockreference to append the attributes</param>
        /// <param name="strAttributeTag">The tag to insert the <paramref name="strAttributeText"/></param>
        /// <param name="strAttributeText">The textstring for <paramref name="strAttributeTag"/></param>
        public static void InsertBlockAttibuteRef(BlockReference blkRef, string strAttributeTag, string strAttributeText)
        {
            Database dbCurrent = HostApplicationServices.WorkingDatabase;
            Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = dbCurrent.TransactionManager;
            using (Transaction tr = tm.StartTransaction())
            {
                BlockTableRecord btAttRec = (BlockTableRecord)tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead);
                foreach (ObjectId idAtt in btAttRec)
                {
                    Entity ent = (Entity)tr.GetObject(idAtt, OpenMode.ForRead);
                    if (ent is AttributeDefinition)
                    {
                        AttributeDefinition attDef = (AttributeDefinition)ent;
                        AttributeReference attRef = new AttributeReference();
                        attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform);
                        if (attRef.Tag == strAttributeTag)
                        {
                            attRef.TextString = strAttributeText;
                            ObjectId idTemp = blkRef.AttributeCollection.AppendAttribute(attRef);
                            tr.AddNewlyCreatedDBObject(attRef, true);
                        }
                    }
                }
                tr.Commit();
            }
        }

        [CommandMethod("plins")]
        public void cmdPLineInsert()
        {
            //Вставляет блоки в вершины полилинии
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;
            TypedValue[] values = new TypedValue[] { new TypedValue(0, "LWPOLYLINE") };
            SelectionFilter filter = new SelectionFilter(values);
            PromptSelectionOptions opts = new PromptSelectionOptions();

            //opts.MessageForRemoval = "\nMust be a type of Block!";
            opts.MessageForAdding = "\nSelect a polyline: ";
            //opts.PrepareOptionalDetails = true;
            //opts.SingleOnly = true;
            //opts.SinglePickInSpace = true;
            //opts.AllowDuplicates = false;
            PromptSelectionResult res = default(PromptSelectionResult);
            res = ed.GetSelection(opts, filter);
            if (res.Status != PromptStatus.OK)
                return;
            //Запрос типа блока
            PromptStringOptions opt = new PromptStringOptions("Введите имя блока: ");
            opt.AllowSpaces = true;
            PromptResult bres = ed.GetString(opt);
            if (bres.Status != PromptStatus.OK) return;
            string blockName = bres.StringResult;
            
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                    if (!acBlkTbl.Has(blockName))
                    {
                        ed.WriteMessage("Блок \"" + blockName + "\" не найден.");
                        tr.Commit();
                        return;
                    }
                    
                    BlockTableRecord space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acBlkTbl[blockName], OpenMode.ForRead);
                    foreach (SelectedObject selPl in res.Value)
                    {
                        Polyline pline = tr.GetObject(selPl.ObjectId, OpenMode.ForRead, false) as Polyline;
                        for (int i = 0; i < pline.NumberOfVertices; i++)
                        {
                            Point2d pnt1 = pline.GetPoint2dAt(i);

                            BlockReference br = new BlockReference(new Point3d(pnt1.X, pnt1.Y, 0), btr.ObjectId);
                            space.AppendEntity(br);
                            tr.AddNewlyCreatedDBObject(br, true);

                            Dictionary<ObjectId, AttInfo> attInfo = new Dictionary<ObjectId, AttInfo>();
                            if (btr.HasAttributeDefinitions)
                            {
                                foreach (ObjectId id in btr)
                                {
                                    DBObject obj =
                                      tr.GetObject(id, OpenMode.ForRead);
                                    AttributeDefinition ad =
                                      obj as AttributeDefinition;

                                    if (ad != null && !ad.Constant)
                                    {
                                        AttributeReference ar =
                                          new AttributeReference();

                                        ar.SetAttributeFromBlock(ad, br.BlockTransform);
                                        ar.Position =
                                          ad.Position.TransformBy(br.BlockTransform);

                                        if (ad.Justify != AttachmentPoint.BaseLeft)
                                        {
                                            ar.AlignmentPoint =
                                              ad.AlignmentPoint.TransformBy(br.BlockTransform);
                                        }
                                        if (ar.IsMTextAttribute)
                                        {
                                            ar.UpdateMTextAttribute();
                                        }

                                        ar.TextString = ad.TextString;

                                        ObjectId arId =
                                          br.AttributeCollection.AppendAttribute(ar);
                                        tr.AddNewlyCreatedDBObject(ar, true);

                                        // Initialize our dictionary with the ObjectId of
                                        // the attribute reference + attribute definition info

                                        attInfo.Add(
                                          arId,
                                          new AttInfo(
                                            ad.Position,
                                            ad.AlignmentPoint,
                                            ad.Justify != AttachmentPoint.BaseLeft
                                          )
                                        );
                                    }
                                }
                            }
                        }

                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
            }
        }

        [CommandMethod("exprt", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void cmdExportExcel()
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;
            PromptSelectionOptions opt = new PromptSelectionOptions();
            PromptSelectionResult res;
            res = ed.GetSelection(opt);
            if (res.Status != PromptStatus.OK) return;
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord acBlkTblRec = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                    List<string> values = new List<string>();
                    //double sum = 0;
                    string s = "";
                    foreach (SelectedObject selPl in res.Value)
                    {
                        DBObject obj = tr.GetObject(selPl.ObjectId, OpenMode.ForRead, false);
                        if (obj.GetType() == typeof(Autodesk.AutoCAD.DatabaseServices.MText))
                        {
                            MText mtxt = obj as MText;
                            s = mtxt.Text;                            
                            values.Add(s);
                            //sum += Convert.ToDouble(s);
                        }
                        if (obj.GetType() == typeof(Autodesk.AutoCAD.DatabaseServices.DBText))
                        {
                            DBText txt = obj as DBText;
                            s = txt.TextString;                           
                            values.Add(s);
                            //sum += Convert.ToDouble(s);
                        }
                        if (obj is Autodesk.AutoCAD.DatabaseServices.Dimension)
                        {
                            Dimension dim = obj as Dimension;
                            if (dim.DimensionText == "")
                            {
                                //sum += dim.Measurement;
                                int dimdec = dim.Dimdec;
                                double val = Math.Round(dim.Measurement, dimdec);
                                values.Add(val.ToString());
                            }
                            else
                            {
                                double val = GetDoubleFString(dim.DimensionText);
                                //sum += Convert.ToDouble(s);
                                values.Add(val.ToString());
                            }
                        }

                    }
                    //create excel
                    Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                    //Книга.
                    ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
                    //Таблица.
                    ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
                    int row = 1;
                    foreach (string val in values)
                    {
                        ObjWorkSheet.Cells[row, 1] = val;
                        row++;
                    }                    
                    ObjExcel.Visible = true;
                    ObjExcel.UserControl = true;

                    //Собираем мусор
                    ObjExcel = null;
                    ObjWorkBook = null;
                    ObjWorkSheet = null;
                    GC.Collect();

                    tr.Commit();
                    //ed.WriteMessage(sum.ToString());
                }
            }

            catch (System.Exception ex)
            {
                ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
            }
        }

        [CommandMethod("NumAttr")]
        public void cmdNumAttributes()
        {
            try
            {
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

                PromptStringOptions strOpt = new PromptStringOptions("Введите имя аттрибута: ");
                PromptResult strRes = ed.GetString(strOpt);
                if (!(strRes.Status == PromptStatus.OK)) { ed.WriteMessage("Programm was cancelled"); return; }
                string attName = strRes.StringResult;

                PromptIntegerOptions intOpt = new PromptIntegerOptions("Начальное число: ");
                PromptIntegerResult intRes = ed.GetInteger(intOpt);
                if (!(intRes.Status == PromptStatus.OK)) { ed.WriteMessage("Programm was cancelled"); return; }
                int num = intRes.Value;

                #region SELECTING
                PromptSelectionOptions selOpt = new PromptSelectionOptions();
                selOpt.MessageForAdding = "Выберите блоки: ";
                TypedValue[] values = { new TypedValue((int)DxfCode.Start, "INSERT"), /*new TypedValue((int)DxfCode.BlockName, "DETAIL_W")*/ };
                SelectionFilter sfilter = new SelectionFilter(values);
                selOpt.AllowDuplicates = false;
                PromptSelectionResult sset = ed.GetSelection(selOpt, sfilter);
                if (!(sset.Status == PromptStatus.OK)) { ed.WriteMessage("Programm was cancelled"); return; }
                #endregion SELECTING

                ObjectId[] objIds = sset.Value.GetObjectIds();
                Database dbCurrent = HostApplicationServices.WorkingDatabase;

                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in objIds)
                    {
                        BlockReference bRef = (BlockReference)tr.GetObject(id, OpenMode.ForWrite);
                        AttributeCollection attcol = bRef.AttributeCollection;

                        bool fnd = false;   //найден ли результат
                        foreach (ObjectId att in attcol)
                        {
                            if (fnd) break;
                            AttributeReference atRef = (AttributeReference)tr.GetObject(att, OpenMode.ForWrite);
                            if (atRef.Tag == attName)
                            {
                                atRef.TextString = num.ToString();
                                num++;
                            }

                        }
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            
        }

        [CommandMethod("cd")]
        public void cmdBlockPointInsert()
        {
            //Вставляет блоки геодезические
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase; 
            string blockName = "g5_330";            
            //"61_Отметки высоты поверхности"
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                    if (!acBlkTbl.Has(blockName))
                    {
                        ed.WriteMessage("Блок \"" + blockName + "\" не найден.");
                        tr.Commit();
                        return;
                    }

                    BlockTableRecord space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acBlkTbl[blockName], OpenMode.ForRead);
                    
                    //Get  point                        
                    PromptPointOptions prPntOpt = new PromptPointOptions("\nУкажите точку: ");
                    PromptPointResult prPntRes = ed.GetPoint(prPntOpt);
                    if (prPntRes.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); return; }
                    Point3d insPnt = prPntRes.Value;

                    //Get height
                    PromptDoubleResult dblRes = ed.GetDouble("\nУкажите высоту: ");
                    double height = dblRes.Value;
                    if (dblRes.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); return; }

                    PromptAngleOptions angOp = new PromptAngleOptions("\nУкажите на правление :");
                    angOp.DefaultValue = PntAngle;
                    PromptDoubleResult ang = ed.GetAngle(angOp);                    
                    if (ang.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); return; }
                    PntAngle = ang.Value;
                    BlockReference br = new BlockReference(new Point3d(insPnt.X,insPnt.Y, height), btr.ObjectId);
                    br.Layer = "61_Отметки высоты поверхности";
                    space.AppendEntity(br);
                    tr.AddNewlyCreatedDBObject(br, true);

                    //insert text                    
                    DBText txt = new DBText();
                    txt.TextString = height.ToString("0.00");
                    txt.Position = new Point3d(insPnt.X+0.25, insPnt.Y+0.25, height);
                    txt.Rotation = PntAngle;
                    txt.Layer = "61_Отметки высоты поверхности";
                    space.AppendEntity(txt);
                    tr.AddNewlyCreatedDBObject(txt, true); 
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
            }
        }


        [CommandMethod("PLexp", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void cmdPolyLineExportExcel()
        {
            //Экспортирует в excel координаты вершин полилинии
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;            

            TypedValue[] values = new TypedValue[] { new TypedValue(0, "LWPOLYLINE") };
            SelectionFilter filter = new SelectionFilter(values);
            PromptSelectionOptions opts = new PromptSelectionOptions();

            //opts.MessageForRemoval = "\nMust be a type of Block!";
            opts.MessageForAdding = "\nSelect a polyline: ";
            //opts.PrepareOptionalDetails = true;
            opts.SingleOnly = true;
            opts.SinglePickInSpace = true;
            opts.AllowDuplicates = false;
            PromptSelectionResult res = default(PromptSelectionResult);
            res = ed.GetSelection(opts, filter);
            if (res.Status != PromptStatus.OK)
                return;            
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    List<double> xAr = new List<double>();
                    List<double> yAr = new List<double>();

                    foreach (SelectedObject selPl in res.Value)
                    {
                        Polyline pline = tr.GetObject(selPl.ObjectId, OpenMode.ForRead, false) as Polyline;
                        for (int i = 0; i < pline.NumberOfVertices; i++)
                        {
                            Point2d pnt1 = pline.GetPoint2dAt(i);
                            xAr.Add(pnt1.X);
                            yAr.Add(pnt1.Y);                            
                        }

                    }
                    tr.Commit();
                    //create excel
                    Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                    //Книга.
                    ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
                    //Таблица.
                    ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
                    int row = 1;
                    foreach (double x in xAr)
                    {
                        ObjWorkSheet.Cells[row, 1] = x;
                        ObjWorkSheet.Cells[row, 2] = yAr[row-1];
                        row++;
                    }
                    ObjExcel.Visible = true;
                    ObjExcel.UserControl = true;

                    //Собираем мусор
                    ObjExcel = null;
                    ObjWorkBook = null;
                    ObjWorkSheet = null;
                    GC.Collect();

                    //tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
            }  
        }


        [CommandMethod("VL_CABLE", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
        public void cmdVLCable()
        {
            //Подсчитывает тяжение кабельной линии
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;

            TypedValue[] values = new TypedValue[] { new TypedValue(0, "LWPOLYLINE") };
            SelectionFilter filter = new SelectionFilter(values);
            PromptSelectionOptions opts = new PromptSelectionOptions();

            opts.MessageForAdding = "\nSelect a polyline: ";            
            opts.SingleOnly = true;
            opts.SinglePickInSpace = true;
            opts.AllowDuplicates = false;

            PromptSelectionResult res = default(PromptSelectionResult);
            res = ed.GetSelection(opts, filter);
            if (res.Status != PromptStatus.OK)
                return;
            try
            {     
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject selPl in res.Value)
                    {
                        //Get weight
                        PromptDoubleOptions dblOpt = new PromptDoubleOptions("Введите массу 1 м кабеля в кг:");                        
                        dblOpt.AllowNegative = false;
                        dblOpt.AllowZero = false;
                        PromptDoubleResult dblRes = ed.GetDouble(dblOpt);
                        if (dblRes.Status != PromptStatus.OK) return;
                        double weight = dblRes.Value;

                        //Get miu
                        dblOpt = new PromptDoubleOptions("Введите коэффициент трения:");
                        dblOpt.AllowNegative = false;
                        dblOpt.AllowZero = false;
                        dblRes = ed.GetDouble(dblOpt);
                        if (dblRes.Status != PromptStatus.OK) return;
                        double miu = dblRes.Value;

                        System.Data.DataTable tbl = new System.Data.DataTable();
                        tbl.Columns.Add("X");
                        tbl.Columns.Add("Y");
                        tbl.Columns.Add("Lenghth");
                        tbl.Columns.Add("Angle");
                        tbl.Columns.Add("T");
                        Polyline pline = tr.GetObject(selPl.ObjectId, OpenMode.ForRead, false) as Polyline;

                        Vector2d curVector = new Vector2d();
                        Vector2d prevVector = new Vector2d();
                        Point2d curPnt = new Point2d();
                        Point2d nextPnt = new Point2d();
                        double curDist = 0;
                        double prevDist=0;
                        double angle = 0;
                        double Tiagenie = 0;  

                        for (int i = 0; i < pline.NumberOfVertices-1; i++)
                        {
                            curPnt = pline.GetPoint2dAt(i);
                            nextPnt = pline.GetPoint2dAt(i+1);
                            curDist = nextPnt.GetDistanceTo(curPnt);                            
                            curVector = new Vector2d(nextPnt.X - curPnt.X, nextPnt.Y - curPnt.Y);
                            if (i > 0)
                            {
                                double mult = prevVector.X * curVector.X+prevVector.Y * curVector.Y;
                                double cos = mult / (curDist * prevDist);
                                if (cos > 1) cos = 1;
                                if (cos < -1) cos = -1;
                                double angleRad = Math.Acos(cos);                                
                                angle = angleRad * 180 / Math.PI;
                                Tiagenie = Tiagenie * Math.Pow(Math.E, miu * angleRad);
                            }
                            Tiagenie += 9.81 * weight * miu * curDist;
                            prevVector = curVector;
                            prevDist = curDist;
                            //double dAng = System.Math.Acos();                            
                            tbl.Rows.Add(curPnt.X, curPnt.Y, curDist, angle, Tiagenie);                           
                        }
                        tbl.Rows.Add(nextPnt.X, nextPnt.Y);
                        ed.WriteMessage("Тяжение составит: " + string.Format("{0:f0}", Tiagenie) + " Н");
                    }
                    tr.Commit();                   
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
            }
        }

        static Point2d PolarPoints(Point2d pPt, double dAng, double dDist)
        {
            return new Point2d(pPt.X + dDist * Math.Cos(dAng),
                               pPt.Y + dDist * Math.Sin(dAng));
        }


        public static void AddOrUpdateAppSettings(string key, string value)
        {
            try
            {
                var configFile = ConfigurationManager.OpenExeConfiguration(Assembly.GetExecutingAssembly().Location);
                var settings = configFile.AppSettings.Settings;
                if (settings[key] == null)
                {
                    settings.Add(key, value);
                }
                else
                {
                    settings[key].Value = value;
                }
                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
            }
            catch (ConfigurationErrorsException)
            {
                MessageBox.Show("Error writing app settings");
            }
        }

        public static void GetOrUpdateAppSettings(string key, out string outValue, string defaultValue)
        {
            outValue = "";
            if (!GetAppSettings(key, out outValue))
            {
                AddOrUpdateAppSettings(key, defaultValue);
                outValue = defaultValue;
            }
        }

        public static bool GetAppSettings(string key, out string outValue)
        {
            outValue = "";
            try
            {
                var configFile = ConfigurationManager.OpenExeConfiguration(Assembly.GetExecutingAssembly().Location);
                var settings = configFile.AppSettings.Settings;
                if (settings[key] == null)
                {                    
                    return false;
                }
                else
                {
                    outValue = settings[key].Value;
                    return true;
                }          
            }
            catch (ConfigurationErrorsException)
            {
                MessageBox.Show("Error getting app settings");
                return false;
            }
        }

        /// <summary>
        /// Чертит кривую провиса провода по данным пользователя
        /// </summary>
        [CommandMethod("VL_DRAW_PROVOD")]
        public void cmdVLDrawProvod()
        {
            try
            {
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;                
                Database db = HostApplicationServices.WorkingDatabase;

                #region GET_SETTINGS
                string sHscale;
                GetOrUpdateAppSettings("hscale", out sHscale, "1000");
                double defaultHscale = GetDoubleFString(sHscale);

                string sVscale;
                GetOrUpdateAppSettings("vscale", out sVscale, "100");
                double defaultVscale = GetDoubleFString(sVscale);

                string sgamma;
                GetOrUpdateAppSettings("gamma", out sgamma, "0.100");
                double defaultGamma = GetDoubleFString(sgamma);

                string sSigma;
                GetOrUpdateAppSettings("sigma", out sSigma, "100");
                double defaultsSigma = GetDoubleFString(sSigma);

                string sGab;
                GetOrUpdateAppSettings("gabarit", out sGab, "6");
                double defaultsGab = GetDoubleFString(sGab); 
                #endregion

                //Get scale
                PromptDoubleOptions dblOpt = new PromptDoubleOptions("\nВведите горизонтальный масштаб 1: ");
                dblOpt.DefaultValue = defaultHscale;
                dblOpt.AllowNegative = false;
                dblOpt.AllowZero = false;
                PromptDoubleResult dblRes = ed.GetDouble(dblOpt);
                if (dblRes.Status != PromptStatus.OK) return;
                double hscale = dblRes.Value;
                AddOrUpdateAppSettings("hscale", hscale.ToString());

                dblOpt.Message = "\nВведите вертикальный масштаб 1: ";
                dblOpt.DefaultValue = defaultVscale;
                dblOpt.AllowNegative = false;
                dblOpt.AllowZero = false;
                dblRes = ed.GetDouble(dblOpt);
                if (dblRes.Status != PromptStatus.OK) return;
                double vscale = dblRes.Value;
                AddOrUpdateAppSettings("vscale", vscale.ToString());


                dblOpt.Message = "\nВведите погонную нагрузку Н/(м*мм2): ";
                dblOpt.DefaultValue = defaultGamma;
                dblOpt.AllowNegative = false;
                dblOpt.AllowZero = false;
                dblRes = ed.GetDouble(dblOpt);
                if (dblRes.Status != PromptStatus.OK) return;
                double gamma = dblRes.Value;
                AddOrUpdateAppSettings("gamma", gamma.ToString());

                dblOpt.Message = "\nВведите напряжение в проводе МПа: ";
                dblOpt.DefaultValue = defaultsSigma;
                dblOpt.AllowNegative = false;
                dblOpt.AllowZero = false;
                dblRes = ed.GetDouble(dblOpt);
                if (dblRes.Status != PromptStatus.OK) return;
                double sigma = dblRes.Value;
                AddOrUpdateAppSettings("sigma", sigma.ToString());

                dblOpt.Message = "\nВведите габарит: ";
                dblOpt.DefaultValue = defaultsGab;
                dblOpt.AllowNegative = false;
                dblOpt.AllowZero = false;
                dblRes = ed.GetDouble(dblOpt);
                if (dblRes.Status != PromptStatus.OK) return;
                double hgab = dblRes.Value;
                AddOrUpdateAppSettings("gabarit", hgab.ToString());

                PromptPointOptions prPntOpt = new PromptPointOptions("\nУкажите  первую точку: ");
                PromptPointResult prPntRes = ed.GetPoint(prPntOpt);
                if (prPntRes.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); return; }
                Point3d Pnt1 = prPntRes.Value;

                while (true)
                {
                    //Get next point
                    prPntOpt = new PromptPointOptions("\nУкажите  следующую точку: ");
                    prPntRes = ed.GetPoint(prPntOpt);
                    if (prPntRes.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); return; }
                    Point3d Pnt2 = prPntRes.Value;

                    //Calc vertices
                    double L = (Pnt2.X - Pnt1.X) * hscale / 1000;
                    double deltaH = (Pnt1.Y - Pnt2.Y) * vscale / 1000;

                    const double npoints = 30;
                    double step = L / npoints;

                    Point3dCollection points = new Point3dCollection();
                    Point3dCollection pointsOffset = new Point3dCollection();

                    for (double x = 0; x < L; x += step)
                    {
                        double fx = gamma * x / (2 * sigma) * (L - x) + x * deltaH / L;
                        double PntX = Pnt1.X + x  * 1000 / hscale;
                        double PntY = Pnt1.Y - fx  * 1000 / vscale;
                        Point3d pnt = new Point3d(PntX, PntY, 0).TransformBy(ed.CurrentUserCoordinateSystem);
                        Point3d pntOffset = new Point3d(PntX, PntY - hgab * 1000 / vscale, 0).TransformBy(ed.CurrentUserCoordinateSystem);
                        points.Add(pnt);
                        pointsOffset.Add(pntOffset);
                    }
                    //add last points
                    points.Add(Pnt2.TransformBy(ed.CurrentUserCoordinateSystem));
                    Point3d Offset = new Point3d(Pnt2.X, Pnt2.Y - hgab * 1000 / vscale, 0).TransformBy(ed.CurrentUserCoordinateSystem);
                    pointsOffset.Add(Offset);

                    //Add polylines
                    Polyline2d pl = new Polyline2d(Poly2dType.SimplePoly, points, 0, false, 0, 0, null);
                    Polyline2d pl2 = new Polyline2d(Poly2dType.SimplePoly, pointsOffset, 0, false, 0, 0, null);

                    //pl2.TransformBy(Matrix3d.Displacement(new Vector3d(0, -hgab * 1000 / vscale, 0));


                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        space.AppendEntity(pl);
                        space.AppendEntity(pl2);
                        tr.AddNewlyCreatedDBObject(pl, true);
                        tr.AddNewlyCreatedDBObject(pl2, true);
                        tr.Commit();
                    }
                    //next point step
                    Pnt1 = Pnt2;
                }
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.ToString());
            }
        }

        /// <summary>
        /// Чертит кривую провиса провода по данным пользователя
        /// </summary>
        [CommandMethod("VL_PROVIS")]
        public void cmdVLProvis()
        {
            try
            {
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Database db = HostApplicationServices.WorkingDatabase;
                DrawCatenaryForm frm = new DrawCatenaryForm();
                DialogResult res = Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(frm);
                //if (res != DialogResult.OK) continue;


            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.ToString());
            }
        }

        /// <summary>
        /// Чертит профиль с коммуникациями
        /// </summary>
        [CommandMethod("VL_PROFILE")]
        public void cmdVLProfile()
        {
            try
            {                               
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Database dbCurrent = HostApplicationServices.WorkingDatabase;
                //вынос текста отметки
                double OptOtmOffset = -8;
                //Высота текста отметки
                double OptOtmHeightText = 2.0;
                //Формат текста отметки
                string OptOtmFormatString = "0.00";
                //Высота текста пояснения для коммуникаций
                double OptDescrTextHeigh = 2.5;

                #region GET_SETTINGS
                string sHscale;
                GetOrUpdateAppSettings("hProfileScale", out sHscale, "100");
                double defaultHscale = GetDoubleFString(sHscale);

                string sVscale;
                GetOrUpdateAppSettings("vProfilescale", out sVscale, "100");
                double defaultVscale = GetDoubleFString(sVscale);
                #endregion

                #region UserInput
                //Get scale
                double hscale = 0;
                double vscale = 0;
                {
                    PromptDoubleOptions dblOpt = new PromptDoubleOptions("\nВведите горизонатльный масштаб 1: ");
                    dblOpt.DefaultValue = defaultHscale;
                    dblOpt.AllowNegative = false;
                    dblOpt.AllowZero = false;
                    PromptDoubleResult dblRes = ed.GetDouble(dblOpt);
                    if (dblRes.Status != PromptStatus.OK) return;
                    hscale = dblRes.Value;
                    AddOrUpdateAppSettings("hscale", hscale.ToString());

                    dblOpt.Message = "\nВведите вертикальный масштаб 1: ";
                    dblOpt.DefaultValue = defaultVscale;
                    dblOpt.AllowNegative = false;
                    dblOpt.AllowZero = false;
                    dblRes = ed.GetDouble(dblOpt);
                    if (dblRes.Status != PromptStatus.OK) return;
                    vscale = dblRes.Value;
                    AddOrUpdateAppSettings("vscale", vscale.ToString());
                }
                
                PromptEntityOptions opt = new PromptEntityOptions("Выберите полилинию профиля:\n");
                opt.SetRejectMessage("Выбранный объект - не полилиния!\n");
                opt.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                
                //Выбор полилинии профиля
                PromptEntityResult res = ed.GetEntity(opt);
                if (res.Status != PromptStatus.OK) return;                
                ObjectId id = res.ObjectId;
                
                Polyline pl = null;
                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {                                            
                    Entity entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (entity.GetType() == typeof(Polyline))
                    {
                        pl = (Polyline)entity;                        
                    }
                    tr.Commit();
                }
                //Список всех точек профиля
                List<ProfilePoint> profPnts = new List<ProfilePoint>();
                //Список пересечений
                List<ProfilePoint> profIntersetions = new List<ProfilePoint>();

                //Высота начальной точки
                {
                    PromptDoubleResult dblRes = ed.GetDouble("\nУкажите высоту начальной точки профиля: ");
                    double height = dblRes.Value;
                    if (dblRes.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); return; }
                    ProfilePoint pnt = new ProfilePoint();
                    pnt.Distance = 0;
                    pnt.Height = height;
                    profPnts.Add(pnt);
                }

                //Высота конечной точки
                {
                    PromptDoubleResult dblRes = ed.GetDouble("\nУкажите высоту конечной точки профиля: ");
                    double height = dblRes.Value;
                    if (dblRes.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); return; }
                    ProfilePoint pnt = new ProfilePoint();
                    pnt.Distance = pl.Length;
                    pnt.Height = height;
                    profPnts.Add(pnt);
                }

                //Ввод характерных точек                
                while (true)
                {
                    //Get  point                        
                    PromptPointOptions prPntOpt = new PromptPointOptions("Укажите характерные точки профиля (черные отметки):\n");
                    PromptPointResult prPntRes = ed.GetPoint(prPntOpt);
                    if (prPntRes.Status != PromptStatus.OK) break;
                    Point3d insPnt = prPntRes.Value;

                    //Get height
                    PromptDoubleOptions pdopt = new PromptDoubleOptions("\nУкажите высоту: ");
                    pdopt.AllowNone = false;
                    PromptDoubleResult dblRes = ed.GetDouble(pdopt);
                    double height = dblRes.Value;
                    if (dblRes.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); return; }

                    ProfilePoint pnt = new ProfilePoint();
                    Point3d onPlPnt = pl.GetClosestPointTo(insPnt, false);
                    double perp = onPlPnt.DistanceTo(insPnt);
                    ed.WriteMessage("Указанная точка находится на расстоянии" + perp.ToString()+"от трассы\n");
                    pnt.Distance = pl.GetDistAtPoint(onPlPnt);
                    pnt.Height = height;
                    profPnts.Add(pnt);
                }

                //Ввод коммуникаций                
                while (true)
                {
                    //Get  point                        
                    PromptPointOptions prPntOpt = new PromptPointOptions("Укажите место пересечения:\n");
                    PromptPointResult prPntRes = ed.GetPoint(prPntOpt);
                    if (prPntRes.Status != PromptStatus.OK) break;
                    Point3d insPnt = prPntRes.Value;

                    double height = 0;
                    double diameter = 0;
                    string descr = "";
                    //Get height
                    {
                        PromptDoubleOptions pdopt = new PromptDoubleOptions("\nУкажите высоту верха коммуникации: ");
                        pdopt.AllowNone = false;
                        PromptDoubleResult dblRes = ed.GetDouble(pdopt);
                        height = dblRes.Value;
                        if (dblRes.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); return; }
                    }
                    //Get diameter
                    {
                        PromptDoubleOptions pdopt = new PromptDoubleOptions("\nУкажите диаметр коммуникации: ");
                        pdopt.AllowNone = false;
                        PromptDoubleResult dblRes = ed.GetDouble(pdopt);
                        diameter = dblRes.Value;
                        if (dblRes.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); return; }
                    }
                    //Get descriptions
                    {
                        PromptStringOptions opt1 = new PromptStringOptions("Введите пояснение: ");
                        opt1.AllowSpaces = true;
                        PromptResult bres = ed.GetString(opt1);
                        if (bres.Status != PromptStatus.OK) return;
                        descr = bres.StringResult;
                    }
                    //
                    ProfilePoint pnt = new ProfilePoint();
                    Point3d onPlPnt = pl.GetClosestPointTo(insPnt, false);
                    double perp = onPlPnt.DistanceTo(insPnt);
                    //ed.WriteMessage("Указанная точка находится на расстоянии" + perp.ToString() + "от трассы\n");
                    pnt.Distance = pl.GetDistAtPoint(onPlPnt);
                    pnt.Height = height;
                    pnt.Diameter = diameter;
                    pnt.Decription = descr;
                    profIntersetions.Add(pnt);
                }

                //Точка вставки профиля 
                Point3d usrPnt;
                {
                    PromptPointOptions prPntOpt = new PromptPointOptions("Укажите точку вставки профиля:\n");
                    PromptPointResult prPntRes = ed.GetPoint(prPntOpt);
                    if (prPntRes.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); return; }
                    usrPnt = prPntRes.Value;
                }
                double BaseHeight = 0;
                //Базовая отметка
                {
                    PromptDoubleOptions pdopt = new PromptDoubleOptions("\nВведите базовую отметку: ");
                    double maxBase = 0;
                    double rezerv = 5;
                    double minporf = profPnts.Min(a => a.Height);
                    double minint = minporf;
                    if (profIntersetions.Count > 0) minint = profIntersetions.Min(a => a.Height);
                    maxBase = Math.Min(minporf, minint) - rezerv;                  

                    pdopt.DefaultValue = maxBase;
                    pdopt.AllowNone = false;
                    PromptDoubleResult dblRes = ed.GetDouble(pdopt);
                    BaseHeight = dblRes.Value;
                    if (dblRes.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); return; }
                }
                #endregion

                #region SORT
                List<ProfilePoint> sortedProfileList = profPnts.OrderBy(x => x.Distance).ToList(); 
                #endregion

                #region DRAW
                
                //линия земли
                Point3dCollection plPnts = new Point3dCollection();
                foreach (ProfilePoint pnt in sortedProfileList)
                {
                    plPnts.Add(new Point3d(usrPnt.X + pnt.Distance * 1000 / hscale, usrPnt.Y + (pnt.Height - BaseHeight) * 1000 / vscale, 0));
                    Point3dCollection vertPnts = new Point3dCollection();
                    //вертикальные линии
                    vertPnts.Add(new Point3d(usrPnt.X + pnt.Distance * 1000 / hscale, usrPnt.Y, 0));
                    vertPnts.Add(new Point3d(usrPnt.X + pnt.Distance * 1000 / hscale, usrPnt.Y + (pnt.Height - BaseHeight) * 1000 / vscale, 0));
                    Polyline2d plVert = new Polyline2d(Poly2dType.SimplePoly, vertPnts, 0, false, 0, 0, null);
                    //отметки
                    DBText hText = new DBText();
                    hText.Position = new Point3d(usrPnt.X + pnt.Distance * 1000 / hscale, usrPnt.Y + OptOtmOffset, 0);
                    hText.Rotation = Math.PI / 2;
                    hText.Justify = AttachmentPoint.MiddleLeft;
                    hText.TextString = pnt.Height.ToString(OptOtmFormatString);
                    hText.AlignmentPoint = hText.Position;
                    hText.Height = OptOtmHeightText;
                    using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord space = (BlockTableRecord)tr.GetObject(dbCurrent.CurrentSpaceId, OpenMode.ForWrite);
                        space.AppendEntity(plVert);
                        space.AppendEntity(hText);
                        tr.AddNewlyCreatedDBObject(plVert, true);
                        tr.AddNewlyCreatedDBObject(hText, true);
                        tr.Commit();
                    }
                }
                //линия земли
                Polyline2d plProfile = new Polyline2d(Poly2dType.SimplePoly, plPnts, 0, false, 0, 0, null);
                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    BlockTableRecord space = (BlockTableRecord)tr.GetObject(dbCurrent.CurrentSpaceId, OpenMode.ForWrite);
                    space.AppendEntity(plProfile);                    
                    tr.AddNewlyCreatedDBObject(plProfile, true);                    
                    tr.Commit();
                }
                //горизонтальная линия
                Point3dCollection H1plPnts = new Point3dCollection();
                H1plPnts.Add(new Point3d(usrPnt.X, usrPnt.Y, 0));
                H1plPnts.Add(new Point3d(usrPnt.X + sortedProfileList.Last().Distance * 1000 / hscale, usrPnt.Y, 0));
                Polyline2d H1pl = new Polyline2d(Poly2dType.SimplePoly, H1plPnts, 0, false, 0, 0, null);
                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    BlockTableRecord space = (BlockTableRecord)tr.GetObject(dbCurrent.CurrentSpaceId, OpenMode.ForWrite);
                    space.AppendEntity(H1pl);
                    tr.AddNewlyCreatedDBObject(H1pl, true);
                    tr.Commit();
                }
                //пересечения
                if (profIntersetions.Count>0)
                {
                    List<ProfilePoint> sortedIntersectionsList = profIntersetions.OrderBy(x => x.Distance).ToList();
                    foreach (ProfilePoint pntIntersect in sortedIntersectionsList)
                    {
                        double xLen = pntIntersect.Diameter * 1000 / hscale;
                        double yLen = pntIntersect.Diameter * 1000 / vscale;
                        double x = usrPnt.X + pntIntersect.Distance * 1000 / hscale;
                        double y = usrPnt.Y + (pntIntersect.Height - pntIntersect.Diameter / 2 - BaseHeight) * 1000 / vscale;
                        double vectorX = 0;
                        double vectorY = 0;
                        double ratio = 0;
                        if (yLen >= xLen)
                        {
                            vectorY = yLen / 2;
                            ratio = xLen / yLen;
                        }
                        else
                        {
                            vectorX = xLen / 2;
                            ratio = yLen / xLen;
                        }
                        Vector3d majAxVec = new Vector3d(vectorX, vectorY, 0);
                        Ellipse elipseObj = new Ellipse(new Point3d(x, y, 0), Vector3d.ZAxis, majAxVec, ratio, 0, 0);
                        Point3dCollection vertPnts = new Point3dCollection();
                        vertPnts.Add(new Point3d(x, usrPnt.Y, 0));
                        vertPnts.Add(new Point3d(x, usrPnt.Y + (pntIntersect.Height - pntIntersect.Diameter - BaseHeight) * 1000 / vscale, 0));
                        Polyline2d plVertIntersect = new Polyline2d(Poly2dType.SimplePoly, vertPnts, 0, false, 0, 0, null);
                        //описание
                        DBText DescrDbText = new DBText();
                        DescrDbText.Position = new Point3d(x, usrPnt.Y, 0);
                        DescrDbText.Rotation = Math.PI / 2;
                        DescrDbText.Justify = AttachmentPoint.BottomLeft;
                        DescrDbText.TextString = pntIntersect.Decription + " d=" + (pntIntersect.Diameter * 1000).ToString("0") + " мм"
                            + " hв=" + pntIntersect.Height.ToString(OptOtmFormatString);
                        DescrDbText.AlignmentPoint = DescrDbText.Position;
                        DescrDbText.Height = OptDescrTextHeigh;
                        
                        using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord space = (BlockTableRecord)tr.GetObject(dbCurrent.CurrentSpaceId, OpenMode.ForWrite);
                            space.AppendEntity(elipseObj);
                            space.AppendEntity(plVertIntersect);
                            space.AppendEntity(DescrDbText);
                            tr.AddNewlyCreatedDBObject(elipseObj, true);
                            tr.AddNewlyCreatedDBObject(plVertIntersect, true);
                            tr.AddNewlyCreatedDBObject(DescrDbText, true);
                            tr.Commit();
                        }
                    }
                }
                #endregion

            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.ToString());
            }
        }

        [CommandMethod("VL_PICKET")]
        public void cmdVLPicket()
        {
            try
            {
                //global variables need to read from options of command!
                int piket_round = 2;
                int offset_round = 2;
                string sAttPicketName;
                string sAttKilometerName;
                string sAttOffsetName;
                string sPicketRound;
                string sAttDistName;                
                string sOffsetRound;
                string sAttOffsetSideName;
                string sAttOffsetLenghtName;
                string sDefaultKeyword;
                #region GET_SETTINGS               
                GetOrUpdateAppSettings("AttPicketName", out sAttPicketName, "ПИКЕТ");
                GetOrUpdateAppSettings("AttKilometerName", out sAttKilometerName, "КИЛОМЕТР");
                GetOrUpdateAppSettings("AttOffsetName", out sAttOffsetName, "СМЕЩЕНИЕ");
                GetOrUpdateAppSettings("AttDistanceName", out sAttDistName, "ДИСТАНЦИЯ");
                GetOrUpdateAppSettings("PicketRound", out sPicketRound, "2");               
                GetOrUpdateAppSettings("OffsetRound", out sOffsetRound, "2");
                GetOrUpdateAppSettings("OffseSideName", out sAttOffsetSideName, "СТОРОНА_СМЕЩЕНИЯ");
                GetOrUpdateAppSettings("sAttOffsetLenghtName", out sAttOffsetLenghtName, "ДИСТАНЦИЯ_СМЕЩЕНИЯ");
                GetOrUpdateAppSettings("sDefaultKeyword", out sDefaultKeyword, "Аттрибут");

                if (int.TryParse(sPicketRound, out piket_round))
                {
                    if(piket_round<0)
                    {
                        piket_round = 0;
                        AddOrUpdateAppSettings("PicketRound", "0");
                    }
                }
                else
                {
                    //set defaults
                    piket_round = 2;
                    AddOrUpdateAppSettings("PicketRound", "2");
                }

                if (int.TryParse(sOffsetRound, out offset_round))
                {
                    if (offset_round < 0)
                    {
                        offset_round = 0;
                        AddOrUpdateAppSettings("OffsetRound", "0");
                    }
                }
                else
                {
                    //set defaults
                    offset_round = 2;
                    AddOrUpdateAppSettings("OffsetRound", "2");
                }
                #endregion

                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Database db = HostApplicationServices.WorkingDatabase;
                //check current space
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                    if(space.Name.ToUpper()!=BlockTableRecord.ModelSpace)
                    {
                        ed.WriteMessage("Программа может быть выполнена только в простанстве модели!\n");
                        return;
                    }
                    tr.Commit();
                }

                //Задана ли текущая трасса?
                VLFileOptions FileOpt = new VLFileOptions();
                if (!FileOpt.TryToRead())
                {
                    if (!SetTrace()) { ed.WriteMessage("\nПрограмма отменена пользователем."); return; }
                    FileOpt.TryToRead();
                }
                //GetKeyword
                PromptKeywordOptions PKWOptions = new PromptKeywordOptions("\nВыберите способ вывода или настройки команды:");
                PKWOptions.Keywords.Add("Аттрибут");
                PKWOptions.Keywords.Add("Текст");
                PKWOptions.Keywords.Add("Мультивыноска");
                PKWOptions.Keywords.Add("Настройки");
                foreach(Keyword kw in PKWOptions.Keywords)
                {
                    if (kw.GlobalName == sDefaultKeyword)
                    {
                        PKWOptions.Keywords.Default = sDefaultKeyword;
                        break;
                    }
                }
                //PKWOptions.Keywords.Default = "Аттрибут";
                PromptResult KwRes = ed.GetKeywords(PKWOptions);
                if (KwRes.Status != PromptStatus.OK) return;

                if (KwRes.StringResult == "Аттрибут")
                {                                        
                    PromptSelectionOptions selOpt = new PromptSelectionOptions();
                    selOpt.MessageForAdding = "Выберите блоки: ";
                    TypedValue[] values = { new TypedValue((int)DxfCode.Start, "INSERT"), /*new TypedValue((int)DxfCode.BlockName, "DETAIL_W")*/ };
                    SelectionFilter sfilter = new SelectionFilter(values);
                    selOpt.AllowDuplicates = false;
                    PromptSelectionResult sset = ed.GetSelection(selOpt, sfilter);
                    if (!(sset.Status == PromptStatus.OK)) { ed.WriteMessage("Программа отменена пользователем.\n"); return; }
                    ObjectId[] objIds = sset.Value.GetObjectIds();
                    long cnt = 0;
                    Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                    using (DocumentLock acLckDoc = doc.LockDocument())
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            foreach (ObjectId id in objIds)
                            {
                                bool isChanged = false;
                                BlockReference bRef = (BlockReference)tr.GetObject(id, OpenMode.ForWrite,true);
                                Point3d insPoint = bRef.Position;
                                //Point3d PntWCS = insPoint.TransformBy(ed.CurrentUserCoordinateSystem);
                                //Point3d pnt2D = new Point3d(PntWCS.X, PntWCS.Y, 0);
                                if (!FileOpt.TryToRead()) return;
                                VLPicketClass picketObj = new VLPicketClass(FileOpt);
                                picketObj.Calc(insPoint, piket_round, offset_round);

                                AttributeCollection attcol = bRef.AttributeCollection;
                                foreach (ObjectId att in attcol)
                                {
                                    AttributeReference atRef = (AttributeReference)tr.GetObject(att, OpenMode.ForWrite, true);
                                    if (atRef.Tag == sAttPicketName)
                                    {
                                        atRef.TextString = picketObj._sPicket;
                                        isChanged = true;
                                    }
                                    if (atRef.Tag == sAttKilometerName)
                                    {
                                        atRef.TextString = picketObj._sKilometer;
                                        isChanged = true;
                                    }

                                    //sAttOffsetSideName
                                    if (atRef.Tag == sAttOffsetSideName)
                                    {
                                        if (picketObj._side == ESide.Middle) atRef.TextString = "";
                                        if (picketObj._side == ESide.Left) atRef.TextString = "-";
                                        if (picketObj._side == ESide.Right) atRef.TextString = "+";
                                        isChanged = true;
                                    }
                                    //sAttOffsetLenghtName
                                    if (atRef.Tag == sAttOffsetLenghtName)
                                    {
                                        atRef.TextString = picketObj._sOffset;
                                    }

                                    if (atRef.Tag == sAttOffsetName)
                                    {
                                        if (picketObj._sSide == "") atRef.TextString = "";
                                        else atRef.TextString = picketObj._sSide + ":" + picketObj._sOffset + " м";
                                        isChanged = true;
                                    }
                                    if (atRef.Tag == sAttDistName)
                                    {
                                        atRef.TextString = picketObj.Dist;
                                        isChanged = true;
                                    }
                                }//foreach attribute
                                if (isChanged) cnt++;
                            }//foreach object
                            tr.Commit();
                        }//using 
                    }//end lock
                    ed.WriteMessage("\nОбновлено " + cnt.ToString() + " блоков\n");
                }
                if (KwRes.StringResult == "Текст")
                {
                    //Get  point                        
                    PromptPointOptions prPntOpt = new PromptPointOptions("\nУкажите точку расчета пикета: ");
                    PromptPointResult prPntRes = ed.GetPoint(prPntOpt);
                    if (prPntRes.Status != PromptStatus.OK) { ed.WriteMessage("\nПрограмма отменена пользователем."); return; }
                    Point3d startPt = prPntRes.Value;
                    //Transform to WCS and set z coordinate t0 null                    
                    Point3d StartPntWCS = startPt.TransformBy(ed.CurrentUserCoordinateSystem);
                    if (!FileOpt.TryToRead()) return;
                    VLPicketClass picketObj = new VLPicketClass(FileOpt);
                    picketObj.Calc(StartPntWCS, piket_round, offset_round);
                    string expr = "префикс+[ПК]+суффикс";
                    ed.WriteMessage(picketObj.Interpret(expr));

                    //Get  point2                        
                    PromptPointOptions prTXTPntOpt = new PromptPointOptions("\nУкажите точку размещения текста: ");
                    prTXTPntOpt.BasePoint = startPt;
                    prTXTPntOpt.UseBasePoint = true;
                    PromptPointResult prTXTPntRes = ed.GetPoint(prTXTPntOpt);
                    if (prTXTPntRes.Status != PromptStatus.OK) { ed.WriteMessage("\nПрограмма отменена пользователем."); return; }
                    Point3d txtPt = prTXTPntRes.Value;
                    Point3d EndPntWCS = txtPt.TransformBy(ed.CurrentUserCoordinateSystem);
                }
                if (KwRes.StringResult == "Мультивыноска")
                {
                    //Get  point                        
                    PromptPointOptions prPntOpt = new PromptPointOptions("\nУкажите исходную точку: ");
                    PromptPointResult prPntRes = ed.GetPoint(prPntOpt);
                    if (prPntRes.Status != PromptStatus.OK) { ed.WriteMessage("\nПрограмма отменена пользователем."); return; }
                    Point3d startPt = prPntRes.Value;
                    //Transform to WCS and set z coordinate t0 null                    
                    Point3d StartPntWCS = startPt.TransformBy(ed.CurrentUserCoordinateSystem);
                    if (!FileOpt.TryToRead()) return;
                    VLPicketClass picketObj = new VLPicketClass(FileOpt);
                    picketObj.Calc(StartPntWCS, piket_round, offset_round);

                    //Get  point2                        
                    PromptPointOptions prMLPntOpt = new PromptPointOptions("\nУкажите точку размещения мультивыноски: ");
                    prMLPntOpt.BasePoint = startPt;
                    prMLPntOpt.UseBasePoint = true;
                    PromptPointResult prMLPntRes = ed.GetPoint(prMLPntOpt);
                    if (prMLPntRes.Status != PromptStatus.OK) { ed.WriteMessage("\nПрограмма отменена пользователем."); return; }
                    Point3d endPt = prMLPntRes.Value;
                    Point3d EndPntWCS = endPt.TransformBy(ed.CurrentUserCoordinateSystem);

                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        BlockTable table = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord model = trans.GetObject(table[BlockTableRecord.ModelSpace],OpenMode.ForWrite) as BlockTableRecord;                        
                        MLeader mld = new MLeader();                        
                        mld.SetDatabaseDefaults();
                        int ldNum = mld.AddLeader();
                        int lnNum = mld.AddLeaderLine(ldNum);
                        mld.AddFirstVertex(lnNum, startPt);
                        mld.AddLastVertex(lnNum, endPt);
                        mld.ContentType = ContentType.MTextContent;
                        MText mText = new MText();
                        mText.SetDatabaseDefaults();
                        mText.Contents = picketObj._sPicket + "\n" + picketObj._sKilometer;                        
                        mText.Location = endPt;
                        mld.MText = mText;
                        mld.TransformBy(ed.CurrentUserCoordinateSystem);
                        model.AppendEntity(mld);
                        trans.AddNewlyCreatedDBObject(mld, true);
                        trans.Commit();
                    }
                }
                if (KwRes.StringResult == "Настройки")
                {
                    PromptKeywordOptions PKWOptions2 = new PromptKeywordOptions("\nВыберите настройки:");
                    PKWOptions2.Keywords.Add("Выбрать трассу");
                    PKWOptions2.Keywords.Add("Задать начальный пикет");
                    PKWOptions2.Keywords.Add("Отслеживание");
                    PKWOptions2.Keywords.Add("Показать трассу");
                    PKWOptions2.Keywords.Default = "Выбрать трассу";
                    PromptResult KwRes2 = ed.GetKeywords(PKWOptions2);
                    if (KwRes2.Status != PromptStatus.OK) return;
                    switch (KwRes2.StringResult)
                    {
                        case "Выбрать":
                            if (!SetTrace()) { ed.WriteMessage("\nПрограмма отменена пользователем."); return; } 
                            break;
                        case "Задать":
                            PromptDoubleOptions dblOpt = new PromptDoubleOptions("\nВведите начальный пикет трассы в метрах:");
                            dblOpt.DefaultValue = 0;
                            PromptDoubleResult dblRes = ed.GetDouble(dblOpt);
                            if (dblRes.Status != PromptStatus.OK) return;
                            FileOpt.profileStartPicket = dblRes.Value;

                            dblOpt.Message = "\nВведите начальный километр трассы в метрах:";
                            dblRes = ed.GetDouble(dblOpt);
                            if (dblRes.Status != PromptStatus.OK) return;
                            FileOpt.profileStartKilometer = dblRes.Value;
                            FileOpt.Write();
                            ed.WriteMessage("\nНастройки трассы установлены.");
                            break;

                        case "Отслеживание":

                            PromptKeywordOptions PKWOptions3 = new PromptKeywordOptions("\nВключить отслеживание пикетов:");
                            PKWOptions3.Keywords.Add("Вкл");
                            PKWOptions3.Keywords.Add("Откл");                           
                            PKWOptions3.Keywords.Default = "Вкл";
                            PromptResult KwRes3 = ed.GetKeywords(PKWOptions3);
                            if (KwRes3.Status != PromptStatus.OK) return;

                            if (KwRes3.StringResult == "Вкл")
                            {
                                FileOpt.VL_PROFILE_MONITOR = true;
                                //this.picketFrm.Visible = true;
                                ed.WriteMessage("\nОтслеживание пикетов включено.");
                            }
                            if (KwRes3.StringResult == "Откл")
                            {
                                FileOpt.VL_PROFILE_MONITOR = false;
                                //this.picketFrm.Visible = false;
                                ed.WriteMessage("\nОтслеживание пикетов выключено.");
                            }
                            FileOpt.Write();
                            break;
                        case "Показать":
                            if (!FileOpt.TryToRead()) return;
                            Handle hndl = FileOpt.profileTraceHandle;
                            using (Transaction trans = db.TransactionManager.StartTransaction())
                            {
                                ObjectId oid = new ObjectId();
                                if (!db.TryGetObjectId(hndl, out oid)) return;
                                ViewEntityPos(oid);
                                trans.Commit();
                            } // using
                            break;
                    }
                }
                AddOrUpdateAppSettings("sDefaultKeyword", KwRes.StringResult);
                //Get  point                        
                //PromptPointOptions prPntOpt = new PromptPointOptions("\nУкажите точку: ");               
                //PromptPointResult prPntRes = ed.GetPoint(prPntOpt);


                //    if (prPntRes.Status != PromptStatus.OK) { ed.WriteMessage("\nПрограмма отменена пользователем."); return; }
                //    Point3d pnt = prPntRes.Value;
                //    //Transform to WCS and set z coordinate t0 null                    
                //    Point3d PntWCS = pnt.TransformBy(ed.CurrentUserCoordinateSystem);
                //    Point3d pnt2D = new Point3d(PntWCS.X, PntWCS.Y, 0);

                //    if (!FileOpt.TryToRead()) return;
                //    Handle hndl = FileOpt.profileTraceHandle;
                //    using (Transaction trans = db.TransactionManager.StartTransaction())
                //    {
                //        ObjectId id = new ObjectId();
                //        if (!db.TryGetObjectId(hndl, out id)) return;
                //        Polyline pline = trans.GetObject(id, OpenMode.ForWrite, false) as Polyline;
                //        if (pline == null) return;
                //        if (pline.Elevation != 0) pline.Elevation = 0;
                //        Point3d onPlPnt = pline.GetClosestPointTo(pnt2D, false);
                //        double picket = pline.GetDistAtPoint(onPlPnt) + FileOpt.profileStartPicket;
                //        double kilometer = pline.GetDistAtPoint(onPlPnt) + FileOpt.profileStartKilometer;
                //        var numPicket = System.Math.Truncate(picket / 100);
                //        var addPicket = picket - numPicket * 100;
                //        string sPicket = String.Format("ПК {0}+{1:f2}", numPicket, addPicket);
                //        var numKM = System.Math.Truncate(kilometer / 1000);
                //        var addKM = kilometer - numKM * 1000;
                //        string sKM = String.Format("КМ {0}+{1:f2}", numKM, addKM);

                //        double perp = onPlPnt.DistanceTo(pnt2D);
                //        string smesh = String.Format("{0:f2}", perp);

                //        string outText = "\n" + sPicket + "\n" + sKM + "\n" + String.Format("Смещение: {0:f2}", perp) + " \n";
                //        ed.WriteMessage(outText);

                //output
                //PromptEntityOptions opt = new PromptEntityOptions("\nВыберите блок: ");
                //opt.SetRejectMessage("\nВыбран не блок!");
                //opt.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.BlockReference), false);
                //PromptEntityResult res = ed.GetEntity(opt);
                //if (res.Status != PromptStatus.OK) return;
                //ObjectId blkId = res.ObjectId;

                //BlockReference bRef = (BlockReference)trans.GetObject(blkId, OpenMode.ForRead);
                //AttributeCollection attcol = bRef.AttributeCollection;
                //foreach (ObjectId att in attcol)
                //{
                //    AttributeReference atRef = (AttributeReference)trans.GetObject(att, OpenMode.ForWrite);

                //    if (atRef.Tag == "ПИКЕТ") atRef.TextString = sPicket;
                //    if (atRef.Tag == "КИЛОМЕТР") atRef.TextString = sKM;
                //    if (atRef.Tag == "СМЕЩЕНИЕ") atRef.TextString = smesh;
                //}                        
                //trans.Commit();
                //} // using                
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.ToString());
            }
        }

        public bool SetTrace()
        {
            try
            {
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Database dbCurrent = HostApplicationServices.WorkingDatabase;
                
                //Выбор пользователем трассы
                PromptEntityOptions opt = new PromptEntityOptions("\nВыберите полилинию трассы:");
                opt.SetRejectMessage("\nВыбранный объект - не полилиния!");
                opt.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);

                //Выбор полилинии
                PromptEntityResult res = ed.GetEntity(opt);
                if (res.Status != PromptStatus.OK) return false;
                ObjectId id = res.ObjectId;
               
                Polyline pl = null;
                VLFileOptions VLopt = new VLFileOptions();

                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    Entity entity = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                    if (entity.GetType() == typeof(Polyline))
                    {
                        pl = (Polyline)entity;
                        //set elevation to 0
                        pl.Elevation = 0;
                    }
                    tr.Commit();
                }
                VLopt.profileTraceHandle = pl.Handle;

                PromptDoubleOptions dblOpt = new PromptDoubleOptions("\nВведите начальный пикет трассы в метрах:");
                dblOpt.DefaultValue = 0;
                PromptDoubleResult dblRes = ed.GetDouble(dblOpt);
                if (dblRes.Status != PromptStatus.OK) return false;
                VLopt.profileStartPicket = dblRes.Value;

                dblOpt.Message = "\nВведите начальный километр трассы в метрах:";
                dblRes = ed.GetDouble(dblOpt);
                if (dblRes.Status != PromptStatus.OK) return false;
                VLopt.profileStartKilometer = dblRes.Value;

                PromptKeywordOptions PKWOptions = new PromptKeywordOptions("\nВключить отслеживание пикетов:");
                PKWOptions.Keywords.Add("Вкл");
                PKWOptions.Keywords.Add("Откл");
                //string[] skwrds = new string[] { "Да", "Нет" };
                //PromptResult KwRes = ed.GetKeywords("\nВключить отслеживание пикетов:", skwrds);
                PKWOptions.Keywords.Default = "Вкл";
                PromptResult KwRes = ed.GetKeywords(PKWOptions);
                if (KwRes.Status != PromptStatus.OK) return false;

                if (KwRes.StringResult == "Вкл")
                {
                    VLopt.VL_PROFILE_MONITOR = true;
                }
                if (KwRes.StringResult == "Откл")
                {
                    VLopt.VL_PROFILE_MONITOR = false;
                }
                VLopt.Write();

                ed.WriteMessage("\nТекущая трасса установлена.");

                return true;
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.ToString());
                return false;
            }
        }

        [CommandMethod("pldiv")]
        public void cmdPLineDivide()
        {
            //Делит полилинии
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;
            TypedValue[] values = new TypedValue[] { new TypedValue(0, "LWPOLYLINE") };
            SelectionFilter filter = new SelectionFilter(values);
            PromptSelectionOptions opts = new PromptSelectionOptions();

            //opts.MessageForRemoval = "\nMust be a type of Block!";
            opts.MessageForAdding = "\nВыберите полилинию: ";
            //opts.PrepareOptionalDetails = true;
            //opts.SingleOnly = true;
            //opts.SinglePickInSpace = true;
            //opts.AllowDuplicates = false;
            PromptSelectionResult res = default(PromptSelectionResult);
            res = ed.GetSelection(opts, filter);
            if (res.Status != PromptStatus.OK) return;
            //Get step
            PromptDoubleOptions dblOpt = new PromptDoubleOptions("Введите шаг:");
            dblOpt.DefaultValue = 30;
            dblOpt.AllowNegative = false;
            dblOpt.AllowZero = false;
            PromptDoubleResult dblRes = ed.GetDouble(dblOpt);
            if (dblRes.Status != PromptStatus.OK) return;
            double step = dblRes.Value;

            //Запрос типа блока
            //PromptStringOptions opt = new PromptStringOptions("Введите имя блока: ");
            //opt.AllowSpaces = true;
            //PromptResult bres = ed.GetString(opt);
            //if (bres.Status != PromptStatus.OK) return;
            //string blockName = bres.StringResult;

            //Get method
            string[] skwrds = new string[] { "1", "2" };
            PromptResult KwRes = ed.GetKeywords("Метод деления: 1-вершины полилинии, 2-точки AutoCAD", skwrds);
            if (KwRes.Status != PromptStatus.OK) return;
            
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    //BlockTable acBlkTbl = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                    //if (!acBlkTbl.Has(blockName))
                    //{
                    //    ed.WriteMessage("Блок \"" + blockName + "\" не найден.");
                    //    tr.Commit();
                    //    return;
                    //}

                    BlockTableRecord space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    //BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acBlkTbl[blockName], OpenMode.ForRead);
                    foreach (SelectedObject selPl in res.Value)
                    {
                        Polyline pline = tr.GetObject(selPl.ObjectId, OpenMode.ForWrite, false) as Polyline;
                        Point2dCollection points = new Point2dCollection();

                        for (int i = 0; i < pline.NumberOfVertices; i++)
                        {
                            Point2d point = pline.GetPoint2dAt(i);
                            points.Add(point);
                        }

                        Point2d prevPnt = points[0];
                        Point2d curPnt = new Point2d();
                        int plVertInx=1;
                        for (int i = 1; i < points.Count; i++)
                        {
                            curPnt = points[i];
                            double dist = prevPnt.GetDistanceTo(curPnt);
                            int ncols = Convert.ToInt32(Math.Ceiling(dist / step));
                            double resStep = dist / ncols;
                            double dAng = prevPnt.GetVectorTo(curPnt).Angle;
                            //Промежуточные точки
                            for (int j = 0; j < ncols-1; j++)
                            {
                                Point2d newPnt = PolarPoints(prevPnt, dAng, resStep);

                                if (KwRes.StringResult == "1")
                                {
                                    pline.AddVertexAt(plVertInx, newPnt, 0, 0, 0);
                                    plVertInx++;
                                }

                                if (KwRes.StringResult == "2")
                                {
                                    DBPoint acPoint = new DBPoint(new Point3d(newPnt.X, newPnt.Y, 0));
                                    acPoint.SetDatabaseDefaults();
                                    space.AppendEntity(acPoint);
                                    tr.AddNewlyCreatedDBObject(acPoint, true);
                                }

                                prevPnt = newPnt;
                            }
                            prevPnt = curPnt;
                            plVertInx++;
                            //BlockReference br = new BlockReference(new Point3d(pnt1.X, pnt1.Y, 0), btr.ObjectId);
                            //space.AppendEntity(br);
                            //tr.AddNewlyCreatedDBObject(br, true);


                        }

                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
            }
        }

        private ObjectId GetObjIDFromHandle(long handle)
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                
                Handle hn = new Handle(handle);
                Database db = HostApplicationServices.WorkingDatabase;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    ObjectId id = db.GetObjectId(false, hn, 0);
                    return id;
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("Exception: " + ex.ToString());
                return ObjectId.Null;
            }
        }

        [CommandMethod("get_handle")]
        public void GetHandle()
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            //PromptEntityOptions opt = new PromptEntityOptions();
            PromptEntityResult res = ed.GetEntity("Select ent:");
            if(res.Status==PromptStatus.OK)
            {
                ObjectId id = res.ObjectId;
                Database dbCurrent = HostApplicationServices.WorkingDatabase;
                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    BlockTableRecord space = (BlockTableRecord)tr.GetObject(dbCurrent.CurrentSpaceId, OpenMode.ForRead);
                    DBObject ob = tr.GetObject(id, OpenMode.ForRead);
                   
                    Handle h = ob.Handle;
                    ed.WriteMessage(h.ToString());
                }
            }
        }

        [CommandMethod("esp_tp")]
        public void cmdEdvansTP()
        {
            try
            {

                #region SELECTING
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                PromptSelectionOptions selOpt = new PromptSelectionOptions();
                selOpt.MessageForAdding = "Выберите блоки абонентов: ";
                TypedValue[] values = { new TypedValue((int)DxfCode.Start, "INSERT"), /*new TypedValue((int)DxfCode.BlockName, "DETAIL_W")*/ };
                SelectionFilter sfilter = new SelectionFilter(values);
                selOpt.AllowDuplicates = false;
                PromptSelectionResult sset = ed.GetSelection(selOpt, sfilter);
                if (!(sset.Status == PromptStatus.OK)) { ed.WriteMessage("Программа отменена пользователем.\n"); return; }

                ObjectId[] objIds = sset.Value.GetObjectIds();
                Database dbCurrent = HostApplicationServices.WorkingDatabase;
                #endregion
                #region FILL_TABLE


                List<Abonent> abonents = new List<Abonent>();
                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in objIds)
                    {
                        BlockReference bRef = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                        AttributeCollection attcol = bRef.AttributeCollection;

                        bool fnd = false;
                        Abonent ab = new Abonent();
                        foreach (ObjectId att in attcol)
                        {
                            AttributeReference atRef = (AttributeReference)tr.GetObject(att, OpenMode.ForRead);

                            if (atRef.Tag == "ТП")
                            {
                                if (atRef.TextString == "")
                                {
                                    string errstr = "В блоке абонента " + " " + ab.NUMBER + " " + ab.FIO + " x=" + bRef.Position.X + " y=" + bRef.Position.Y + " не указана ТП!";
                                    MessageBox.Show(errstr, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                                ab.TP = atRef.TextString;
                                fnd = true;
                            }
                            if (atRef.Tag == "НОМЕР_ЗАЯВКИ")
                            {
                                ab.NUMBER = atRef.TextString;
                            }
                            if (atRef.Tag == "ФИО")
                            {
                                ab.FIO = atRef.TextString;
                            }
                            if (atRef.Tag == "МОЩНОСТЬ")
                            {
                                try
                                {
                                    ab.POWER = GetDoubleFString(atRef.TextString);
                                }
                                catch (System.Exception)
                                {
                                    ab.POWER = 0;
                                    string errstr = "В блоке абонента " + " " + ab.NUMBER + " " + ab.FIO + " x=" + bRef.Position.X + " y=" + bRef.Position.Y + " не указана мощность!";
                                    MessageBox.Show(errstr, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            if (atRef.Tag == "ФИДЕР")
                            {
                                ab.FIDER = atRef.TextString;
                            }
                            if (atRef.Tag == "ФИДЕР_МАРКА_ПРОВОДА")
                            {
                                ab.PROVOD = atRef.TextString;
                            }
                            if (atRef.Tag == "ФИДЕР_ДЛИНА_ЛИНИИ")
                            {
                                ab.LENGHT = atRef.TextString;
                            }

                        }
                        if (fnd) abonents.Add(ab);
                    }
                }
                #endregion FILL_TABLE

                #region SORT_TABLE

                var TPAbGroups = from ab in abonents
                                 group ab by ab.TP;

                //Get method
                string[] skwrds = new string[] { "1", "2" };
                PromptResult KwRes = ed.GetKeywords("Метод вывода: 1-лист AutoCAD, 2-лист Excel", skwrds);
                if (KwRes.Status != PromptStatus.OK) return;

                if (KwRes.StringResult == "1")
                {
                    //add and fill acad tables in layouts
                    #region FILL_LAYOUT
                    int cntr = 1;
                    foreach (IGrouping<string, Abonent> g in TPAbGroups)
                    {
                        Transaction trans = dbCurrent.TransactionManager.StartTransaction();
                        {
                            try
                            {
                                string tpName = g.Key;

                                //copy layout
                                Layout lt = null;
                                string curName = "";
                                if (cntr == 1)
                                {
                                    lt = ImportLayoutWithOrWithoutReplace(_TemplatePath + "template.dwg", "Таблица ТП");
                                    curName = "Таблица ТП";
                                }
                                else
                                {
                                    lt = ImportLayoutWithOrWithoutReplace(_TemplatePath + "template.dwg", "Таблица ТП л.2");
                                    curName = "Таблица ТП л.2";
                                }
                                cntr++;

                                if (lt == null)
                                {
                                    ed.WriteMessage("Лист не удалось скопировать!");
                                    return;
                                }

                                //rename layout                          
                                LayoutManager.Current.RenameLayout(curName, tpName);
                                //lt.LayoutName = tpName;

                                //Get tables
                                BlockTableRecord layoutTableRecord = (BlockTableRecord)lt.Database.TransactionManager.GetObject(lt.BlockTableRecordId, OpenMode.ForWrite, false);

                                Table tblSmall = null;
                                Table tblFiders = null;

                                foreach (ObjectId entityId in layoutTableRecord)
                                {
                                    object acadObject = lt.Database.TransactionManager.
                                        GetObject(entityId, OpenMode.ForRead, false);
                                    Table ACADtbl = acadObject as Table;
                                    if (ACADtbl != null)
                                    {
                                        //check what it is?
                                        if (ACADtbl.Cells[0, 0].TextString == "Номер ТП")
                                        {
                                            tblSmall = lt.Database.TransactionManager.
                                        GetObject(entityId, OpenMode.ForWrite, false) as Table;
                                        }
                                        if (ACADtbl.Cells[0, 0].TextString == "Номер группы, линии (фидер)")
                                        {
                                            tblFiders = lt.Database.TransactionManager.
                                        GetObject(entityId, OpenMode.ForWrite, false) as Table;
                                        }

                                    }
                                }

                                //ObjectId tblID = GetFirstTable(lt);
                                //Table tbl = lt.Database.TransactionManager.GetObject(tblID, OpenMode.ForWrite, false) as Table;


                                tblSmall.Cells[0, 1].TextString = tpName;

                                List<Abonent> TPlist = new List<Abonent>();
                                foreach (var t in g)
                                {
                                    //MessageBox.Show(t.FIO);
                                    //longAbons += t.FIO = "; ";
                                    TPlist.Add(t);
                                }

                                int fidercol = 1;
                                var TPfiederGroups = from ab in TPlist
                                                     orderby ab.FIDER
                                                     group ab by ab.FIDER;
                                //TPfiederGroups.OrderBy(FIDER);
                                double SumPower = 0;
                                foreach (IGrouping<string, Abonent> f in TPfiederGroups)
                                {
                                    string longAbons = "";
                                    double power = 0;
                                    //MessageBox.Show(f.Key);
                                    foreach (var a in f)
                                    {
                                        longAbons += a.FIO + "; ";
                                        power += a.POWER;
                                        //MessageBox.Show(a.FIO);

                                        //f.Sum(a.POWER);
                                    }
                                    tblFiders.Cells[0, fidercol].TextString = f.Key;//fider
                                    tblFiders.Cells[1, fidercol].TextString = power.ToString("0.00");//fider
                                    tblFiders.Cells[2, fidercol].TextString = longAbons;
                                    tblFiders.Cells[3, fidercol].TextString = f.Count().ToString();//count abonenst on fieder
                                    if (f.ElementAt(0).PROVOD != "") tblFiders.Cells[4, fidercol].TextString = f.ElementAt(0).PROVOD;
                                    if (f.ElementAt(0).LENGHT != "") tblFiders.Cells[10, fidercol].TextString = f.ElementAt(0).LENGHT;

                                    fidercol++;
                                    SumPower += power;
                                }
                                tblSmall.Cells[1, 1].TextString = SumPower.ToString("0.00");
                            }
                            finally
                            {
                                trans.Commit();
                            }
                        }
                        //MessageBox.Show(g.Key);
                    }
                    ed.Regen();
                    #endregion
                }
                #region Excel
                if (KwRes.StringResult == "2")
                {
                    //Get Cntr
                    PromptIntegerOptions opt = new PromptIntegerOptions("Введите номер первого листа:");
                    opt.AllowNegative = false;
                    opt.AllowZero = false;
                    opt.DefaultValue = 1;
                    PromptIntegerResult res = ed.GetInteger(opt);
                    if(res.Status!=PromptStatus.OK)
                    {
                        ed.WriteMessage("Прервано пользователем");
                        return;
                    }
                    int cntr = res.Value;
                    //create excel
                    Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                    try
                    {
                        //Get excel file
                        string xlsTmplFile = _TemplatePath + "Однолинейная схема- расчет.xlsx";
                        if (!File.Exists(xlsTmplFile))
                        {
                            MessageBox.Show("Файл " + xlsTmplFile + " не найден");
                            return;
                        }
                        string curDwgPath = ed.Document.Name;
                        string workFolder = Path.GetDirectoryName(curDwgPath);

 

                        
                        foreach (IGrouping<string, Abonent> g in TPAbGroups)
                        {
                            string tpName = g.Key;
                            //copy file
                            string newXlsFile = workFolder + "\\" + cntr.ToString() +". " + tpName + ". Однолинейная схема- расчет.xlsx";

                            //if file exist?
                            if (File.Exists(newXlsFile))
                            {
                                DialogResult resInt = MessageBox.Show("Файл " + newXlsFile + " существует. Заменить?", "vl_tools", MessageBoxButtons.YesNo);
                                if (resInt == DialogResult.Yes)
                                {
                                    //replace file;
                                    File.Copy(xlsTmplFile, newXlsFile, true);
                                }
                                else
                                {
                                    //next step
                                    cntr++;
                                    continue;
                                }
                            }
                            else
                            {
                                //copy no perlace
                                File.Copy(xlsTmplFile, newXlsFile, false);
                            }

                            //Книга.
                            ObjWorkBook = ObjExcel.Workbooks.Open(newXlsFile);
                            //Таблица.
                            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets["Расчет1"];

                            //Fill sheet
                            //Dwg Properties
                            DatabaseSummaryInfo dbInfo = dbCurrent.SummaryInfo;
                            DatabaseSummaryInfoBuilder dbInfoBldr = new DatabaseSummaryInfoBuilder(dbInfo); 
                            
                            ObjWorkSheet.Range["R41"].Value = dbInfoBldr.Title;
                            ObjWorkSheet.Range["R43"].Value = dbInfoBldr.Comments;
                            ObjWorkSheet.Range["R46"].Value = dbInfoBldr.Subject;
                            ObjWorkSheet.Range["M46"].Value = dbInfoBldr.Author;
                            ObjWorkSheet.Range["M51"].Value = dbInfoBldr.Keywords;
                            ObjWorkSheet.Range["X47"].Value = cntr.ToString() + ".1";
                            ObjWorkSheet.Range["Z97"].Value = cntr.ToString() + ".2";

                            
                            ObjWorkSheet.Range["R49"].Value = "Однолинейная принципиальная схема РУ-10 / 0,4кВ " + tpName;
                            ObjWorkSheet.Range["U4"].Value = tpName;

                            List<Abonent> TPlist = new List<Abonent>();
                            foreach (var t in g)
                            {
                                //MessageBox.Show(t.FIO);
                                //longAbons += t.FIO = "; ";
                                TPlist.Add(t);
                            }

                            int fidercol = 1;
                            var TPfiederGroups = from ab in TPlist
                                                 orderby ab.FIDER
                                                 group ab by ab.FIDER;                            
                           
                            foreach (IGrouping<string, Abonent> f in TPfiederGroups)
                            {
                                string longAbons = "";
                                double power = 0;
                                //MessageBox.Show(f.Key);
                                foreach (var a in f)
                                {
                                    longAbons += a.FIO + "; ";
                                    power += a.POWER;                                   
                                }
                                string fiderName = f.Key;
                                //double fiderNum = this.GetDoubleFString(fiderName);
                                switch (fiderName)
                                {
                                    case "Л-1":
                                        {
                                            ObjWorkSheet.Range["H20"].Value = longAbons;
                                            ObjWorkSheet.Range["H26"].Value = f.Count();
                                            ObjWorkSheet.Range["H31"].Value = power;
                                            if (f.ElementAt(0).PROVOD != "")
                                            {
                                                ObjWorkSheet.Range["H37"].Value = f.ElementAt(0).PROVOD;
                                            }
                                            if (f.ElementAt(0).LENGHT != "")
                                            {
                                                ObjWorkSheet.Range["H30"].Value = f.ElementAt(0).LENGHT;
                                            }
                                            break;
                                        }
                                    case "Л-2":
                                            ObjWorkSheet.Range["I20"].Value = longAbons;
                                            ObjWorkSheet.Range["I26"].Value = f.Count();
                                            ObjWorkSheet.Range["I31"].Value = power;
                                            if (f.ElementAt(0).PROVOD != "")
                                            {
                                                ObjWorkSheet.Range["I37"].Value = f.ElementAt(0).PROVOD;
                                            }
                                            if (f.ElementAt(0).LENGHT != "")
                                            {
                                                ObjWorkSheet.Range["I30"].Value = f.ElementAt(0).LENGHT;
                                            }
                                            break;
                                    case "Л-3":
                                            ObjWorkSheet.Range["J20"].Value = longAbons;
                                            ObjWorkSheet.Range["J26"].Value = f.Count();
                                            ObjWorkSheet.Range["J31"].Value = power;
                                            if (f.ElementAt(0).PROVOD != "")
                                            {
                                                ObjWorkSheet.Range["J37"].Value = f.ElementAt(0).PROVOD;
                                            }
                                            if (f.ElementAt(0).LENGHT != "")
                                            {
                                                ObjWorkSheet.Range["J30"].Value = f.ElementAt(0).LENGHT;
                                            }
                                            break;
                                    case "Л-4":
                                            ObjWorkSheet.Range["K20"].Value = longAbons;
                                            ObjWorkSheet.Range["K26"].Value = f.Count();
                                            ObjWorkSheet.Range["K31"].Value = power;
                                            if (f.ElementAt(0).PROVOD != "")
                                            {
                                                ObjWorkSheet.Range["K37"].Value = f.ElementAt(0).PROVOD;
                                            }
                                            if (f.ElementAt(0).LENGHT != "")
                                            {
                                                ObjWorkSheet.Range["K30"].Value = f.ElementAt(0).LENGHT;
                                            }
                                            break;
                                    case "Л-5":
                                            ObjWorkSheet.Range["Q20"].Value = longAbons;
                                            ObjWorkSheet.Range["Q26"].Value = f.Count();
                                            ObjWorkSheet.Range["Q31"].Value = power;
                                            if (f.ElementAt(0).PROVOD != "")
                                            {
                                                ObjWorkSheet.Range["Q37"].Value = f.ElementAt(0).PROVOD;
                                            }
                                            if (f.ElementAt(0).LENGHT != "")
                                            {
                                                ObjWorkSheet.Range["Q30"].Value = f.ElementAt(0).LENGHT;
                                            }
                                            break;
                                    case "Л-6":
                                            ObjWorkSheet.Range["T20"].Value = longAbons;
                                            ObjWorkSheet.Range["T26"].Value = f.Count();
                                            ObjWorkSheet.Range["T31"].Value = power;
                                            if (f.ElementAt(0).PROVOD != "")
                                            {
                                                ObjWorkSheet.Range["T37"].Value = f.ElementAt(0).PROVOD;
                                            }
                                            if (f.ElementAt(0).LENGHT != "")
                                            {
                                                ObjWorkSheet.Range["T30"].Value = f.ElementAt(0).LENGHT;
                                            }
                                            break;
                                    case "Л-7":
                                            ObjWorkSheet.Range["W20"].Value = longAbons;
                                            ObjWorkSheet.Range["W26"].Value = f.Count();
                                            ObjWorkSheet.Range["W31"].Value = power;
                                            if (f.ElementAt(0).PROVOD != "")
                                            {
                                                ObjWorkSheet.Range["W37"].Value = f.ElementAt(0).PROVOD;
                                            }
                                            if (f.ElementAt(0).LENGHT != "")
                                            {
                                                ObjWorkSheet.Range["W30"].Value = f.ElementAt(0).LENGHT;
                                            }
                                            break;
                                }
                                //switch(
                                //tblFiders.Cells[0, fidercol].TextString = f.Key;//fider
                                //tblFiders.Cells[1, fidercol].TextString = power.ToString("0.00");//fider
                                //tblFiders.Cells[2, fidercol].TextString = longAbons;
                                //tblFiders.Cells[3, fidercol].TextString = f.Count().ToString();//count abonenst on fieder
                                //if (f.ElementAt(0).PROVOD != "") tblFiders.Cells[4, fidercol].TextString = f.ElementAt(0).PROVOD;
                                //if (f.ElementAt(0).LENGHT != "") tblFiders.Cells[10, fidercol].TextString = f.ElementAt(0).LENGHT;
                                fidercol++;
                                
                            }                            


                            //Save and close Exls
                            ObjWorkBook.Save();
                            ObjWorkBook.Close();
                            //int row = 1;
                            //foreach (string val in values)
                            //{
                            //    ObjWorkSheet.Cells[row, 1] = val;
                            //    row++;
                            //}
                            //ObjExcel.Visible = true;
                            //ObjExcel.UserControl = true;
                            ObjWorkBook = null;
                            ObjWorkSheet = null;
                            cntr++;
                        }
                        //Собираем мусор
                        ObjExcel = null;
                        ObjWorkBook = null;
                        ObjWorkSheet = null;
                        GC.Collect();
                    }
                    catch (System.Exception ex)
                    {
                        //Собираем мусор
                        ObjExcel = null;
                        ObjWorkBook = null;
                        ObjWorkSheet = null;
                        GC.Collect();
                        MessageBox.Show(ex.ToString());
                    }




                #endregion
                #endregion
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        [CommandMethod("esp_shem")]
        public void cmdEdvansShema()
        {
            try
            {

                #region SELECTING
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                PromptSelectionOptions selOpt = new PromptSelectionOptions();
                selOpt.MessageForAdding = "Выберите блоки абонентов: ";
                TypedValue[] values = { new TypedValue((int)DxfCode.Start, "INSERT"), /*new TypedValue((int)DxfCode.BlockName, "DETAIL_W")*/ };
                SelectionFilter sfilter = new SelectionFilter(values);
                selOpt.AllowDuplicates = false;
                PromptSelectionResult sset = ed.GetSelection(selOpt, sfilter);
                if (!(sset.Status == PromptStatus.OK)) { ed.WriteMessage("Программа отменена пользователем.\n"); return; }

                ObjectId[] objIds = sset.Value.GetObjectIds();
                Database dbCurrent = HostApplicationServices.WorkingDatabase;
                #endregion
                #region FILL_TABLE


                List<Abonent> abonents = new List<Abonent>();
                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in objIds)
                    {
                        BlockReference bRef = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                        AttributeCollection attcol = bRef.AttributeCollection;

                        bool fnd = false;
                        Abonent ab = new Abonent();
                        foreach (ObjectId att in attcol)
                        {
                            AttributeReference atRef = (AttributeReference)tr.GetObject(att, OpenMode.ForRead);

                            if (atRef.Tag == "ТП")
                            {
                                if (atRef.TextString == "")
                                {
                                    string errstr = "В блоке абонента " + " " + ab.NUMBER + " " + ab.FIO + " x=" + bRef.Position.X + " y=" + bRef.Position.Y + " не указана ТП!";
                                    MessageBox.Show(errstr, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                                ab.TP = atRef.TextString;
                                fnd = true;
                            }
                            if (atRef.Tag == "НОМЕР_ЗАЯВКИ")
                            {
                                ab.NUMBER = atRef.TextString;
                                if (atRef.TextString == "")
                                {
                                    string errstr = "В блоке абонента " + " " + ab.NUMBER + " " + ab.FIO + " x=" + bRef.Position.X + " y=" + bRef.Position.Y + " не указан номер заявки!";
                                    MessageBox.Show(errstr, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                            }
                            if (atRef.Tag == "ФИО")
                            {
                                ab.FIO = atRef.TextString;
                                if (atRef.TextString == "")
                                {
                                    string errstr = "В блоке абонента " + " " + ab.NUMBER + " " + ab.FIO + " x=" + bRef.Position.X + " y=" + bRef.Position.Y + " не указана фамилия!";
                                    MessageBox.Show(errstr, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                            }
                            if (atRef.Tag == "МОЩНОСТЬ")
                            {
                                try
                                {
                                    ab.POWER = GetDoubleFString(atRef.TextString);
                                }
                                catch (System.Exception)
                                {
                                    ab.POWER = 0;
                                    string errstr = "В блоке абонента " + " " + ab.NUMBER + " " + ab.FIO + " x=" + bRef.Position.X + " y=" + bRef.Position.Y + " не указана мощность!";
                                    MessageBox.Show(errstr, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            if (atRef.Tag == "ФИДЕР")
                            {
                                ab.FIDER = atRef.TextString;
                            }
                            if (atRef.Tag == "ФИДЕР_МАРКА_ПРОВОДА")
                            {
                                ab.PROVOD = atRef.TextString;
                            }
                            if (atRef.Tag == "ФИДЕР_ДЛИНА_ЛИНИИ")
                            {
                                ab.LENGHT = atRef.TextString;
                            }

                        }
                        if (fnd) abonents.Add(ab);
                    }
                }
                #endregion FILL_TABLE               
                
                PromptIntegerOptions options2 = new PromptIntegerOptions("Введите номер первого листа:")
                {
                    AllowNegative = false,
                    AllowZero = false,
                    DefaultValue = 1
                };
                PromptIntegerResult integer = ed.GetInteger(options2);
                if (integer.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("Прервано пользователем");
                    return;
                }
                int num = integer.Value;
                var TPAbGroups = from ab in abonents
                                 group ab by ab.TP;
                    
                //int cntr = 1;
                foreach (IGrouping<string, Abonent> g in TPAbGroups)
                {
                    //Show dlg
                    NewTPForm frm = new NewTPForm();
                    frm.m_group = g;
                    frm.templatePath = _TemplatePath + "ТП";
                    frm.connection = connection;
                    DialogResult TPres = Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(frm);
                    if (TPres != DialogResult.OK) continue;

                    //Получаем ТП и автоматы
                    //DataGridView grid = frm.dataGridViewLines;
                    //DataGridView grid2 = frm.dataGridView2;
                    //List<int> list2 = frm.dgv1Indexes;

                    //List<string> LAutomates = new List<string>();
                    //LAutomates = frm.automates;
                    //string filePath = frm.SelectedFilePath;
                    //string blockName = frm.SelectedBlockName;
                    //bool flag15 = frm.checkBoxIsExistingTP.Checked;

                    //// Create the value GUIDs.
                    //string newBlockName = Guid.NewGuid().ToString();
                    //ImportBlock(filePath, blockName, newBlockName);
                    //BlockReference br = BlockJig(newBlockName, new Hashtable());
                    ////br.Explode(new DBObjectCollection());
                    
                    //Transaction trans = dbCurrent.TransactionManager.StartTransaction();
                    //{
                    //    try
                    //    {
                    //        BlockTable bt = (BlockTable)trans.GetObject(dbCurrent.BlockTableId,
                    //                  OpenMode.ForRead, false);
                    //        BlockTableRecord record = (BlockTableRecord)trans.GetObject(br.BlockTableRecord,
                    //            OpenMode.ForWrite, false);
                    //        Table TblFiders = null;
                    //        Table TblElems = null;
                    //        //Опросник
                    //        Table TblOprosnik = null;
                    //        Table TblTP = null;

                    //        foreach (ObjectId entityId in record)
                    //        {
                    //            DBObject acadObject = trans.GetObject(entityId, OpenMode.ForRead, false);
                    //            BlockReference reference4 = acadObject as BlockReference;
                    //            if (reference4 != null)
                    //            {
                    //                AttributeCollection attributes2 = reference4.AttributeCollection;
                    //                bool flag18 = false;
                    //                foreach (ObjectId id4 in attributes2)
                    //                {
                    //                    AttributeReference reference5 = (AttributeReference)trans.GetObject(id4, OpenMode.ForWrite);
                    //                    string str8 = "";
                    //                    if (grid2.Rows[0].Cells[1].Value != null)
                    //                    {
                    //                        str8 = grid2.Rows[0].Cells[1].Value.ToString();
                    //                    }
                    //                    if (reference5.Tag == "НАЗВАНИЕ_ЧЕРТЕЖА")
                    //                    {
                    //                        reference5.TextString = "Схема однолинейная принципиальная сети 0,4 кВ (" + str8 + ")";
                    //                        flag18 = true;
                    //                    }
                    //                }
                    //                foreach (ObjectId id5 in attributes2)
                    //                {
                    //                    AttributeReference reference6 = (AttributeReference)trans.GetObject(id5, OpenMode.ForWrite);
                    //                    if (reference6.Tag == "ЛИСТ")
                    //                    {
                    //                        if (flag18)
                    //                        {
                    //                            if (flag15)
                    //                            {
                    //                                reference6.TextString = num.ToString();
                    //                            }
                    //                            else
                    //                            {
                    //                                reference6.TextString = num.ToString() + ".1";
                    //                            }
                    //                        }
                    //                        else
                    //                        {
                    //                            reference6.TextString = num.ToString() + ".2";
                    //                        }
                    //                    }
                    //                }
                    //            }

                    //            Table ACADtbl = acadObject as Table;
                    //            if (ACADtbl != null)
                    //            {
                    //                //check what it is?
                    //                if (ACADtbl.Cells[0, 1].TextString == "Номер группы, линии (фидер)")
                    //                {
                    //                    TblFiders = trans.GetObject(entityId, OpenMode.ForWrite, false) as Table;
                    //                }
                    //                if (ACADtbl.Cells[0, 0].TextString == "Ведомость силового оборудования")
                    //                {
                    //                    TblElems = trans.GetObject(entityId, OpenMode.ForWrite, false) as Table;
                    //                }
                    //                if (ACADtbl.Cells[0, 0].TextString == "Номер ТП")
                    //                {
                    //                    TblTP = trans.GetObject(entityId, OpenMode.ForWrite, false) as Table;
                    //                }
                    //                //Находим опросник
                    //                if (ACADtbl.Cells[0, 0].TextString.Contains("ОПРОСНЫЙ ЛИСТ"))
                    //                {
                    //                    TblOprosnik = trans.GetObject(entityId, OpenMode.ForWrite, false) as Table;
                    //                }

                    //            }
                    //        }
                    //        //Fill tables
                    //        if (TblFiders != null)
                    //        {
                    //            TblFiders.InsertColumns(2, 25.0, LAutomates.Count);
                    //            int col = 2;

                    //            foreach (int index in list2)
                    //            //for (int i = 0; i < LAutomates.Count; i++)
                    //            {
                    //                if (grid.Rows[9].Cells[index].Value != null)
                    //                    TblFiders.Cells[0, col].TextString = grid.Columns[index].Name;
                    //                if (grid.Rows[9].Cells[index].Value != null)
                    //                    TblFiders.Cells[1, col].TextString = grid.Rows[9].Cells[index].Value.ToString();
                    //                if (grid.Rows[8].Cells[index].Value != null)
                    //                    TblFiders.Cells[2, col].TextString = grid.Rows[8].Cells[index].Value.ToString() + " А";
                    //                if (grid.Rows[6].Cells[index].Value != null)
                    //                    TblFiders.Cells[3, col].TextString = grid.Rows[6].Cells[index].Value.ToString();
                    //                if (grid.Rows[7].Cells[index].Value != null)
                    //                    TblFiders.Cells[4, col].TextString = grid.Rows[7].Cells[index].Value.ToString();
                    //                if (grid.Rows[5].Cells[index].Value != null)
                    //                    TblFiders.Cells[5, col].TextString = grid.Rows[5].Cells[index].Value.ToString();
                    //                if (grid.Rows[1].Cells[index].Value != null)
                    //                    TblFiders.Cells[6, col].TextString = grid.Rows[1].Cells[index].Value.ToString();
                    //                if (grid.Rows[0].Cells[index].Value != null)
                    //                    TblFiders.Cells[7, col].TextString = grid.Rows[0].Cells[index].Value.ToString() + " кВт";
                    //                if (grid.Rows[3].Cells[index].Value != null)
                    //                    TblFiders.Cells[8, col].TextString = grid.Rows[3].Cells[index].Value.ToString();
                    //                if (grid.Rows[4].Cells[index].Value != null)
                    //                    TblFiders.Cells[9, col].TextString = grid.Rows[4].Cells[index].Value.ToString() + " кВт";
                    //                if (grid.Rows[2].Cells[index].Value != null)
                    //                    TblFiders.Cells[10, col].TextString = grid.Rows[2].Cells[index].Value.ToString();
                    //                col++;
                    //            }
                    //        }
                    //        if (TblElems != null)
                    //        {
                    //            int count = TblElems.Rows.Count;
                    //            double height = TblElems.Rows[count - 1].Height;
                    //            int i = 1;
                    //            int row = 0;
                    //            foreach (int index in list2)
                    //            {
                    //                if (!((grid.Rows[4].Cells[index].Value == null) & flag15))
                    //                {
                    //                    TblElems.InsertRows((count - 1) + row, height, 1);
                    //                    TblElems.Cells[(count - 1) + row, 0].TextString = "QF" + i.ToString();
                    //                    string marka = "ВА51-35 ";
                    //                    if (Convert.ToDouble(LAutomates[i - 1]) > 250.0)
                    //                    {
                    //                        marka = "ВА57-39 ";
                    //                    }
                    //                    TblElems.Cells[(count - 1) + row, 1].TextString = marka + LAutomates[i - 1] + " А";
                    //                    TblElems.Cells[(count - 1) + row, 1].Alignment = CellAlignment.MiddleLeft;
                    //                    TblElems.Cells[(count - 1) + row, 2].TextString = "1";
                    //                    row++;
                    //                }
                    //                i++;
                    //            }

                    //        }
                    //        if (TblTP != null)
                    //        {
                    //            if (grid2.Rows[0].Cells[1].Value != null)
                    //            {
                    //                TblTP.Cells[0, 1].TextString = grid2.Rows[0].Cells[1].Value.ToString();
                    //            }
                    //            if (grid2.Rows[1].Cells[1].Value != null)
                    //            {
                    //                TblTP.Cells[1, 1].TextString = grid2.Rows[1].Cells[1].Value.ToString();
                    //            }
                    //            if (grid2.Rows[3].Cells[1].Value != null)
                    //            {
                    //                TblTP.Cells[2, 1].TextString = grid2.Rows[3].Cells[1].Value.ToString();
                    //            }
                    //            double Tok = (Convert.ToDouble(grid2.Rows[5].Cells[1].Value) / 1.73) / 0.38;
                    //            TblTP.Cells[3, 1].TextString = string.Format("{0:f2}", Tok) + " А";
                    //        }

                    //        //Заполняем опросник
                    //        if (TblOprosnik != null)
                    //        {
                    //            TblOprosnik.Cells[0, 0].TextString = "ОПРОСНЫЙ ЛИСТ №" + cntr.ToString() +
                    //                " (" + grid2.Rows[0].Cells[1].Value.ToString() + ")" + "\r\n"
                    //                +  "на трансформаторную подстанцию наружной установки";
                    //            List<string> nominals = new List<string>();
                    //            for (int i = 2; i < TblOprosnik.Columns.Count; i++)
                    //            {
                    //                string cellTxt=TblOprosnik.Cells[24, i].TextString;
                    //                if (cellTxt != "") nominals.Add(cellTxt);
                    //            }

                    //            foreach (string autm in LAutomates) if (!nominals.Contains(autm)) 
                    //                MessageBox.Show("Автомат " + autm + " не встречается в опроснике!");

                    //            for (int i = 0; i < nominals.Count(); i++)
                    //            {
                    //                var slctd = from s in LAutomates.AsEnumerable()
                    //                            where s.Equals(nominals[i])
                    //                            select s;
                    //                int scnt = slctd.Count();
                    //                if (scnt > 0) TblOprosnik.Cells[25, i + 2].TextString = scnt.ToString();
                    //            }                                
                    //        }
                    //    }
                    //    finally
                    //    {
                    //        trans.Commit();
                    //    }
                    //    cntr++;
                    //    num++;  
                    //}                   
                }
                ed.Regen();  
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }            
        }


        [CommandMethod("ESP_CHECK")]
        public void cmdCheckAbonentsSQL()
        {
            System.Data.DataTable dbAbonsTable = new System.Data.DataTable();
            System.Data.DataTable srvrAbonsTable = null;
            #region SELECTING
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            PromptSelectionOptions selOpt = new PromptSelectionOptions();
            selOpt.MessageForAdding = "Выберите блоки абонентов: ";
            TypedValue[] values = { new TypedValue((int)DxfCode.Start, "INSERT"), /*new TypedValue((int)DxfCode.BlockName, "DETAIL_W")*/ };
            SelectionFilter sfilter = new SelectionFilter(values);
            selOpt.AllowDuplicates = false;
            PromptSelectionResult sset = ed.GetSelection(selOpt, sfilter);
            if (!(sset.Status == PromptStatus.OK)) { ed.WriteMessage("Программа отменена пользователем.\n"); return; }

            ObjectId[] objIds = sset.Value.GetObjectIds();
            Database dbCurrent = HostApplicationServices.WorkingDatabase;
            #endregion
            try
            {
                #region FILL_TABLE

                List<Abonent> abonents = new List<Abonent>();
                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in objIds)
                    {
                        BlockReference bRef = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                        AttributeCollection attcol = bRef.AttributeCollection;

                        bool fnd = false;
                        Abonent ab = new Abonent();
                        foreach (ObjectId att in attcol)
                        {
                            AttributeReference atRef = (AttributeReference)tr.GetObject(att, OpenMode.ForRead);

                            if (atRef.Tag == "ТП")
                            {
                                if (atRef.TextString == "")
                                {
                                    string errstr = "В блоке абонента " + " " + ab.NUMBER + " " + ab.FIO + " x=" + bRef.Position.X + " y=" + bRef.Position.Y + " не указана ТП!";
                                    MessageBox.Show(errstr, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                                ab.TP = atRef.TextString;
                                fnd = true;
                            }
                            if (atRef.Tag == "НОМЕР_ЗАЯВКИ")
                            {
                                ab.NUMBER = atRef.TextString;
                            }
                            if (atRef.Tag == "ФИО")
                            {
                                ab.FIO = atRef.TextString;
                            }
                            if (atRef.Tag == "МОЩНОСТЬ")
                            {
                                try
                                {
                                    ab.POWER = GetDoubleFString(atRef.TextString);
                                }
                                catch (System.Exception)
                                {
                                    ab.POWER = 0;
                                    string errstr = "В блоке абонента " + " " + ab.NUMBER + " " + ab.FIO + " x=" + bRef.Position.X + " y=" + bRef.Position.Y + " не указана мощность!";
                                    MessageBox.Show(errstr, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            if (atRef.Tag == "ФИДЕР")
                            {
                                ab.FIDER = atRef.TextString;
                            }
                            if (atRef.Tag == "ФИДЕР_МАРКА_ПРОВОДА")
                            {
                                ab.PROVOD = atRef.TextString;
                            }
                            if (atRef.Tag == "ФИДЕР_ДЛИНА_ЛИНИИ")
                            {
                                ab.LENGHT = atRef.TextString;
                            }

                        }
                        if (fnd) abonents.Add(ab);
                    }
                }
                dbAbonsTable.Columns.Add("НОМЕР_ЗАЯВКИ");
                dbAbonsTable.Columns.Add("ФИО");
                foreach (Abonent abonent in abonents)
                {
                    dbAbonsTable.Rows.Add(abonent.NUMBER, abonent.FIO);
                }

                #endregion FILL_TABLE
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            #region GET_SQL
            try
            {
                //Show dlg
                SQLTest1.SQLAbonentsForm frm = new SQLTest1.SQLAbonentsForm();                
                DialogResult res = Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(frm);
                if (res != DialogResult.OK) return;
                if (frm.table == null) return;                
                else srvrAbonsTable = frm.table;
                                
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            #endregion
            #region CHECK
            #endregion
        }

        [CommandMethod("SL_gredit", CommandFlags.UsePickSet)]
        public void cmdSmartLineGroupEdit()
        {
            try
            {
                List<BlockObject> selDwgObjects = new List<BlockObject>();  //список выбранных объектов

                //string appName = "ESMT_LEP_v1.0";
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;               
                PromptSelectionResult selection = ed.SelectImplied();
                if (selection.Status != PromptStatus.OK)
                {

                    TypedValue tv = new TypedValue(0, "INSERT");
                    SelectionFilter sf = new SelectionFilter(new TypedValue[] { tv });                    
                    selection = ed.GetSelection(sf);
                }
               
                if (selection.Status != PromptStatus.OK)
                    return;

                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    try
                    {
                        ObjectId[] ids = selection.Value.GetObjectIds();
                        
                        foreach (ObjectId id in ids)
                        {
                            Entity entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity.GetType() == typeof(BlockReference))
                            {
                                BlockObject dwgObj = new BlockObject(id);
                                if (!dwgObj.HasExtData) continue;
                                selDwgObjects.Add(dwgObj);
                            }

                            //dwgObj.SaveXMLtoCADEntity(dwgObj.ToXElement());

                            /*DBObject obj = tr.GetObject(id, OpenMode.ForRead);
                            ObjectId extensionDictionaryId = obj.ExtensionDictionary;
                            DBDictionary extDictionary = (DBDictionary)tr.GetObject(extensionDictionaryId, OpenMode.ForRead);
                            if (extDictionary.Contains(appName))
                            {
                                ObjectId xrID = extDictionary.GetAt(appName);
                                Xrecord readBack = (Xrecord)tr.GetObject(xrID, OpenMode.ForRead);
                                string val = (string)readBack.Data.AsArray()[0].Value;
                                XDocument XmlDoc = XDocument.Parse(val);

                                //Номер не всегда обновлен! Нужно обновлять по значению аттрибута SL_NUM!
                                string number = XmlDoc.Root.Attribute("number").Value;

                                tblNames.Columns.Add(id.ToString());
                                tblDetailsFirst.Columns.Add(id.ToString());
                                tblDetailsSecond.Columns.Add(id.ToString());
                                col++;

                                tblNames.Rows[0][col] = number;

                                string type = XmlDoc.Root.Attribute("name").Value;
                                tblNames.Rows[1][col] = type;

                                //first specification
                                IEnumerable<XElement> first = from el in XmlDoc.Root.Elements("Specification")
                                                              where (string)el.Attribute("name") == "FirstTable"
                                                              select el;
                                if (first.Count() != 1) return;
                                foreach (XElement el in first.Elements())
                                {
                                    string name = el.Attribute("name").Value;
                                    string count = el.Attribute("count").Value;
                                    if (!detailsFirstTable.Contains(name))
                                    {
                                        detailsFirstTable.Add(name);
                                        DataRow newrow = tblDetailsFirst.NewRow();
                                        tblDetailsFirst.Rows.Add(newrow);
                                    }
                                    int row = detailsFirstTable.IndexOf(name);
                                    tblDetailsFirst.Rows[row][0] = name;
                                    tblDetailsFirst.Rows[row][col] = count;
                                    //MessageBox.Show(el.Attribute("name").Value + "-" + el.Attribute("count").Value);
                                }

                                //second specification
                                IEnumerable<XElement> second = from el in XmlDoc.Root.Elements("Specification")
                                                               where (string)el.Attribute("name") == "SecondTable"
                                                               select el;
                                if (second.Count() != 1) return;
                                foreach (XElement el in second.Elements())
                                {
                                    string name = el.Attribute("name").Value;
                                    string count = el.Attribute("count").Value;
                                    if (!detailsSecondTable.Contains(name))
                                    {
                                        detailsSecondTable.Add(name);
                                        DataRow newrow = tblDetailsSecond.NewRow();
                                        tblDetailsSecond.Rows.Add(newrow);
                                    }
                                    int row = detailsSecondTable.IndexOf(name);
                                    tblDetailsSecond.Rows[row][0] = name;
                                    tblDetailsSecond.Rows[row][col] = count;
                                }//foreach (XElement el in second.Elements()                        
                            }//if (extDictionary.Contains(appName))*/
                        }//foreach (ObjectId id in ids)

                        //Form edit
                        SLGroupEditFrm frm = new SLGroupEditFrm(selDwgObjects);                        

                        DialogResult Dlgres = Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(frm);
                        if (Dlgres != DialogResult.OK) return;



                    }//try
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    finally
                    {
                        tr.Commit();
                    }
                    
                }//using (var tr = doc.TransactionManager.StartTransaction())
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        [CommandMethod("смартобъем", CommandFlags.UsePickSet)]
        public void cmdSmartLineVolumes()
        {
            try
            {
                List<BlockObject> selBlockObjects = new List<BlockObject>();  //список выбранных объектов
                List<PlineObject> selPlineObject = new List<PlineObject>();  //список выбранных объектов
                //string appName = "ESMT_LEP_v1.0";
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;
                PromptSelectionResult selection = ed.SelectImplied();
                if (selection.Status != PromptStatus.OK)
                {

                    TypedValue[] valueArray = new TypedValue[4];
                    valueArray.SetValue(new TypedValue(-4, "<or"), 0);
                    valueArray.SetValue(new TypedValue(0, "INSERT"), 1);
                    valueArray.SetValue(new TypedValue(0, "LWPOLYLINE"), 2);
                    valueArray.SetValue(new TypedValue(-4, "or>"), 3);
                    SelectionFilter filter = new SelectionFilter(valueArray);
                    selection = ed.GetSelection(filter);
                }

                if (selection.Status != PromptStatus.OK)
                    return;

                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    try
                    {
                        ObjectId[] ids = selection.Value.GetObjectIds();

                        foreach (ObjectId id in ids)
                        {
                            Entity entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity.GetType() == typeof(BlockReference))
                            {
                                BlockObject dwgObj = new BlockObject(id);
                                if (!dwgObj.HasExtData) continue;
                                selBlockObjects.Add(dwgObj);
                            }
                            else if (entity.GetType() == typeof(Polyline))
                            {
                                DwgObject dwgObj = new DwgObject(id);
                                if (dwgObj.HasExtData)
                                {
                                    if (dwgObj.ObjectType == "trench")
                                    {
                                        //xElem.Add(TrenchObject.Open(id).ToXSpecification());
                                    }
                                    else
                                    {
                                        PlineObject plobj = PlineObject.Open(id);
                                        selPlineObject.Add(plobj);
                                    }
                                }
                            }                           
                        }
                        //Form edit
                        string TemplatePath = _TemplatePath + "Объемы";
                        VolumeForm frm = new VolumeForm(selBlockObjects, selPlineObject, TemplatePath);
                        DialogResult Dlgres = Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(frm);
                        if (Dlgres != DialogResult.OK) return;

                    }//try
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    finally
                    {
                        tr.Commit();
                    }

                }//using (var tr = doc.TransactionManager.StartTransaction())
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        [CommandMethod("SL_exprtxml")]
        public void cmdSmartLineExportXML()
        {
            try
            {
                //List<BlockObject> selDwgObjects = new List<BlockObject>();  //список выбранных объектов                
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                Database db = doc.Database;

                var tv = new TypedValue(0, "INSERT");
                var sf = new SelectionFilter(new TypedValue[] { tv });
                var res = ed.GetSelection(sf);

                if (res.Status != PromptStatus.OK) return;

                //Get folder
                FolderBrowserDialog dlg = new FolderBrowserDialog();
                DialogResult dres = dlg.ShowDialog();
                if (dres != DialogResult.OK) return;
                string path = dlg.SelectedPath;

                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    try
                    {
                        ObjectId[] ids = res.Value.GetObjectIds();

                        foreach (ObjectId id in ids)
                        {
                            BlockObject dwgObj = new BlockObject(id);
                            if (!dwgObj.HasExtData) continue;
                            XElement el = dwgObj.ToXElement();
                            el.Save(path + "\\" + dwgObj.Name + ".xml");
                        }

                    }//try
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    finally
                    {
                        tr.Commit();
                    }

                }//using (var tr = doc.TransactionManager.StartTransaction())
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        [CommandMethod("testimport")]
        public void TestImport()
        {
            ImportDwg("D:\\ВР32-31.DWG");
        }

        [CommandMethod("vl_tableat")]
        public void AttributeTableEditor()
        {
            #region SELECTING
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            PromptSelectionOptions selOpt = new PromptSelectionOptions();
            selOpt.MessageForAdding = "Выберите блоки: ";
            TypedValue[] values = { new TypedValue((int)DxfCode.Start, "INSERT"), /*new TypedValue((int)DxfCode.BlockName, "DETAIL_W")*/ };
            SelectionFilter sfilter = new SelectionFilter(values);
            selOpt.AllowDuplicates = false;
            PromptSelectionResult sset = ed.GetSelection(selOpt, sfilter);
            if (!(sset.Status == PromptStatus.OK)) { ed.WriteMessage("Программа отменена пользователем.\n"); return; }

            ObjectId[] objIds = sset.Value.GetObjectIds();
            Database dbCurrent = HostApplicationServices.WorkingDatabase;
            #endregion
            try
            {
                #region FILL_TABLES
                List<string> blockNames = new List<string>();

                List<NamedBlockRef> bResList = new List<NamedBlockRef>();
                List<NamedBlockRefsCollection> bRefCollections = new List<NamedBlockRefsCollection>();
    
                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in objIds)
                    {
                        BlockReference bRef = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                        if (bRef != null) 
                        {
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bRef.DynamicBlockTableRecord, OpenMode.ForRead);                            
                            if (btr.HasAttributeDefinitions)
                            {
                                NamedBlockRef nbRef = new NamedBlockRef(bRef, btr.Name); 
                                bResList.Add(nbRef);
                            }
                        }
                    }
                }
                var groups = bResList.AsEnumerable().GroupBy(b => b.Name);
                foreach (var group in groups)
                {
                    NamedBlockRefsCollection col = new NamedBlockRefsCollection();
                    foreach (NamedBlockRef nbRef in group.ToArray())
                    {
                        col.Add(nbRef);
                    }
                    bRefCollections.Add(col);
                    //MessageBox.Show(group.Key + group.Count().ToString());
                }
                #endregion FILL_TABLES
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        [CommandMethod("трассвл")]
        public void cmdVLineTrace()
        {
            #region CONFIG

            string TraceGroup = "";
            string DBPath = _TemplatePath + "ВЛ\\ВЛ.xml";
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;
            PromptSelectionResult res=null;

            try
            {                
                Configuration config = null;
                try
                {
                    config = ConfigurationManager.OpenExeConfiguration(Assembly.GetExecutingAssembly().Location);
                    TraceGroup = config.AppSettings.Settings["TraceGroup"].Value;
                }
                catch (System.Exception)
                {
                    config.AppSettings.Settings.Add("TraceGroup", "");
                    config.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                }
                TraceGroup = config.AppSettings.Settings["TraceGroup"].Value;
                #endregion

                //Вставляет блоки в вершины полилинии
                //Options
                //string curGroup="";
                
                //End Options
 
                TypedValue[] values = new TypedValue[] { new TypedValue(0, "LWPOLYLINE") };
                SelectionFilter filter = new SelectionFilter(values);

                PromptSelectionOptions opts = new PromptSelectionOptions();
                opts.AllowDuplicates = false;

                opts.Keywords.Add("Options");

                string kws = opts.Keywords.GetDisplayString(true);
                opts.MessageForAdding = "\nВыберите трассу ВЛ или " + kws;
                opts.KeywordInput +=
                   delegate (object sender, SelectionTextInputEventArgs e)
                   {
                   //ed.WriteMessage("\nKeyword entered: {0}", e.Input);
                   TraceOptionsFrm frm = new TraceOptionsFrm(DBPath, TraceGroup);
                       DialogResult TraceOptionsRes = Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(frm);
                   //if (res != DialogResult.OK) return;
                   TraceGroup = frm.GroupName;
                   //Save options
                   config.AppSettings.Settings["TraceGroup"].Value = TraceGroup;
                       config.Save(ConfigurationSaveMode.Modified);
                       ConfigurationManager.RefreshSection("appSettings");

                   };
                res = ed.GetSelection(opts, filter);
                if (res.Status == PromptStatus.OK)
                {
                    //ed.WriteMessage(res.ToString());
                }
                else
                {
                    ed.WriteMessage("Прервано пользователем\n");
                    return;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            DataSet ds = new DataSet();
            try
            {
                //Init database frame                
                System.Data.DataTable table = new System.Data.DataTable();
                table.Columns.Add(new System.Data.DataColumn("GROUP_NAME", typeof(string), "", MappingType.Attribute));
                table.Columns.Add(new System.Data.DataColumn("ANGLE", typeof(string), "", MappingType.Attribute));
                table.Columns.Add(new System.Data.DataColumn("TYPE", typeof(string), "", MappingType.Attribute));                
                table.Columns.Add(new System.Data.DataColumn("XMLNAME", typeof(string), "", MappingType.Attribute));
                table.Columns.Add(new System.Data.DataColumn("BLOCKNAME", typeof(string), "", MappingType.Attribute));
                table.Columns.Add(new System.Data.DataColumn("BLOCKSCALE", typeof(string), "", MappingType.Attribute));
                ds.Tables.Add(table);
                //table.Rows.Add("0,4 кВ", 1, "ВЛ.dwg", "04_Пром");
                
                //ds.WriteXml(DBPath);                
                if (!File.Exists(DBPath)) throw new FileNotFoundException("Файл " + DBPath + " не найден!");
                ds.ReadXml(DBPath, XmlReadMode.IgnoreSchema);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
                     
           
            try
            {
                //Set Options
                EnumerableRowCollection<DataRow> curGroupRows;  //Current Group
                var rows = ds.Tables[0].AsEnumerable();
                if (TraceGroup == "") TraceGroup = rows.First().Field<string>("GROUP_NAME");
                curGroupRows = rows.Where(row => row.Field<string>("GROUP_NAME") == TraceGroup);
                if (curGroupRows.Count() == 0) return;
                

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                    BlockTableRecord space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    
                    foreach (SelectedObject selPl in res.Value)
                    {
                        Polyline pline = tr.GetObject(selPl.ObjectId, OpenMode.ForRead, false) as Polyline;
                        Point2d prevPnt, curPnt, nextPnt;                      
                        
                        for (int i = 1; i < pline.NumberOfVertices-1; i++)
                        {
                            prevPnt = pline.GetPoint2dAt(i-1);
                            curPnt = pline.GetPoint2dAt(i);
                            nextPnt = pline.GetPoint2dAt(i+1);
                            Vector2d v1 = prevPnt.GetVectorTo(curPnt);
                            Vector2d v2 = curPnt.GetVectorTo(nextPnt);

                            //double dAng = System.Math.Abs(dAng2 - dAng1);
                            double dAng = v1.GetAngleTo(v2);
                            dAng = dAng * 180 / Math.PI;
                            
                            var BiggestAngles = curGroupRows.Where(row => Convert.ToDouble(row.Field<string>("ANGLE")) >= dAng);
                            var ZazemlenieRow = curGroupRows.First(row => row.Field<string>("TYPE") == "ЗАЗЕМЛЕНИЕ");
                            if (BiggestAngles.Count() > 0)
                            {
                                var arow = BiggestAngles.OrderBy(r => Convert.ToDouble(r.Field<string>("ANGLE"))).First();
                                if (arow != null)
                                {
                                    string xmlName = arow.Field<string>("XMLNAME");
                                    string filePath = _TemplatePath + "ВЛ\\" + xmlName;
                                    string blockName = arow.Field<string>("BLOCKNAME");
                                    double blockScale = Convert.ToDouble(arow.Field<string>("BLOCKSCALE"));
                                    if (!acBlkTbl.Has(blockName))
                                    {
                                        ed.WriteMessage("Указанный блок с именем: " + blockName + " - отсутствует в чертеже!");
                                        continue;
                                    }
                                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acBlkTbl[blockName], OpenMode.ForRead);
                                    BlockReference br = new BlockReference(new Point3d(curPnt.X, curPnt.Y, 0), btr.ObjectId);
                                    br.ScaleFactors = new Scale3d(blockScale, blockScale, blockScale);
                                    space.AppendEntity(br);
                                    tr.AddNewlyCreatedDBObject(br, true);

                                    //Add attributes
                                    Dictionary<ObjectId, AttInfo> attInfo = new Dictionary<ObjectId, AttInfo>();
                                    if (btr.HasAttributeDefinitions)
                                    {
                                        foreach (ObjectId id in btr)
                                        {
                                            DBObject obj =
                                              tr.GetObject(id, OpenMode.ForRead);
                                            AttributeDefinition ad =
                                              obj as AttributeDefinition;

                                            if (ad != null && !ad.Constant)
                                            {
                                                AttributeReference ar =
                                                  new AttributeReference();

                                                ar.SetAttributeFromBlock(ad, br.BlockTransform);
                                                ar.Position =
                                                  ad.Position.TransformBy(br.BlockTransform);

                                                if (ad.Justify != AttachmentPoint.BaseLeft)
                                                {
                                                    ar.AlignmentPoint =
                                                      ad.AlignmentPoint.TransformBy(br.BlockTransform);
                                                }
                                                if (ar.IsMTextAttribute)
                                                {
                                                    ar.UpdateMTextAttribute();
                                                }

                                                ar.TextString = ad.TextString;

                                                ObjectId arId =
                                                  br.AttributeCollection.AppendAttribute(ar);
                                                tr.AddNewlyCreatedDBObject(ar, true);

                                                // Initialize our dictionary with the ObjectId of
                                                // the attribute reference + attribute definition info

                                                attInfo.Add(
                                                  arId,
                                                  new AttInfo(
                                                    ad.Position,
                                                    ad.AlignmentPoint,
                                                    ad.Justify != AttachmentPoint.BaseLeft
                                                  )
                                                );
                                            }
                                        }
                                    }
                                    //End Add attributes
                                    XElement el = XElement.Load(filePath);
                                    BlockObject blObj = new BlockObject(el);
                                    blObj.Number = "0";
                                    blObj.Name = xmlName.Replace(".xml", "");
                                    blObj.ObjId = br.ObjectId;
                                    blObj.SaveXMLtoCADEntity(blObj.ToXElement());
                                    blObj.SetAttributes();
                                }
                            }
                            //Заземление
                            if (ZazemlenieRow!= null)
                            {
                                string xmlName = ZazemlenieRow.Field<string>("XMLNAME");
                                string filePath = _TemplatePath + "ВЛ\\" + xmlName;
                                string blockName = ZazemlenieRow.Field<string>("BLOCKNAME");
                                double blockScale = Convert.ToDouble(ZazemlenieRow.Field<string>("BLOCKSCALE"));
                                if (!acBlkTbl.Has(blockName))
                                {
                                    ed.WriteMessage("Указанный блок с именем: " + blockName + " - отсутствует в чертеже!");
                                    continue;
                                }
                                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acBlkTbl[blockName], OpenMode.ForRead);
                                BlockReference br = new BlockReference(new Point3d(curPnt.X, curPnt.Y, 0), btr.ObjectId);
                                br.ScaleFactors = new Scale3d(blockScale, blockScale, blockScale);
                                space.AppendEntity(br);
                                tr.AddNewlyCreatedDBObject(br, true);

                                //Add attributes
                                Dictionary<ObjectId, AttInfo> attInfo = new Dictionary<ObjectId, AttInfo>();
                                if (btr.HasAttributeDefinitions)
                                {
                                    foreach (ObjectId id in btr)
                                    {
                                        DBObject obj =
                                          tr.GetObject(id, OpenMode.ForRead);
                                        AttributeDefinition ad =
                                          obj as AttributeDefinition;

                                        if (ad != null && !ad.Constant)
                                        {
                                            AttributeReference ar =
                                              new AttributeReference();

                                            ar.SetAttributeFromBlock(ad, br.BlockTransform);
                                            ar.Position =
                                              ad.Position.TransformBy(br.BlockTransform);

                                            if (ad.Justify != AttachmentPoint.BaseLeft)
                                            {
                                                ar.AlignmentPoint =
                                                  ad.AlignmentPoint.TransformBy(br.BlockTransform);
                                            }
                                            if (ar.IsMTextAttribute)
                                            {
                                                ar.UpdateMTextAttribute();
                                            }

                                            ar.TextString = ad.TextString;

                                            ObjectId arId =
                                              br.AttributeCollection.AppendAttribute(ar);
                                            tr.AddNewlyCreatedDBObject(ar, true);

                                            // Initialize our dictionary with the ObjectId of
                                            // the attribute reference + attribute definition info

                                            attInfo.Add(
                                              arId,
                                              new AttInfo(
                                                ad.Position,
                                                ad.AlignmentPoint,
                                                ad.Justify != AttachmentPoint.BaseLeft
                                              )
                                            );
                                        }
                                    }
                                }
                                //End Add attributes
                                XElement el = XElement.Load(filePath);
                                BlockObject blObj = new BlockObject(el);
                                blObj.Number = "0";
                                blObj.Name = xmlName.Replace(".xml", "");
                                blObj.ObjId = br.ObjectId;
                                blObj.SaveXMLtoCADEntity(blObj.ToXElement());
                                blObj.SetAttributes();
                            }

                            //prevPnt = curPnt;
                        }  
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
            }            

        }

        [CommandMethod("VL_BLOCKTOPOLYDIM")]
        public void cmdBlockToPolyDim()
        {
            try
            {
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Database dbCurrent = HostApplicationServices.WorkingDatabase;
                #region GET_SETTINGS                
                string sOffset;
                GetOrUpdateAppSettings("VL_BLOCKTOPOLYDIM_OFFSET", out sOffset, "0");
                double defaultOffset = GetDoubleFString(sOffset);
                #endregion

                #region Выбор полилиний
                TypedValue[] values = new TypedValue[] { new TypedValue(0, "LWPOLYLINE,POLYLINE") };
                SelectionFilter filter = new SelectionFilter(values);
                PromptSelectionOptions opts = new PromptSelectionOptions();
                opts.MessageForAdding = "Выберите полилинии, от которых нужно проставить размеры:\n: ";
                opts.AllowDuplicates = false;
                PromptSelectionResult plselres = ed.GetSelection(opts, filter);
                if (!(plselres.Status == PromptStatus.OK)) { ed.WriteMessage("Программа отменена пользователем.\n"); return; }
                ObjectId[] ids = plselres.Value.GetObjectIds();

                List<Curve> plines = new List<Curve>();
                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId id in ids)
                    {
                        Entity entity = (Entity) tr.GetObject(id, OpenMode.ForRead);
                        Curve crv = (Curve) entity;
                        plines.Add(crv);
                        crv.Highlight();                        
                    }                    
                    tr.Commit();
                }                
                #endregion

                #region Выбор блоков
                PromptSelectionOptions selOpt = new PromptSelectionOptions();
                selOpt.MessageForAdding = "Выберите блоки: ";
                TypedValue tv = new TypedValue(0, "INSERT");
                SelectionFilter sfilter = new SelectionFilter(new TypedValue[] { tv });
                selOpt.AllowDuplicates = false;
                PromptSelectionResult sset = ed.GetSelection(selOpt, sfilter);
                if (!(sset.Status == PromptStatus.OK)) { ed.WriteMessage("Программа отменена пользователем.\n"); return; }
                ObjectId[] blockIds = sset.Value.GetObjectIds();
                #endregion

                #region offset
                PromptDistanceOptions pdOPts = new PromptDistanceOptions("Укажите смещение: ");
                pdOPts.DefaultValue = defaultOffset;
                PromptDoubleResult dblres = ed.GetDistance(pdOPts);
                if (!(dblres.Status == PromptStatus.OK)) { ed.WriteMessage("Программа отменена пользователем.\n"); return; }
                double offset = dblres.Value;
                AddOrUpdateAppSettings("VL_BLOCKTOPOLYDIM_OFFSET", offset.ToString());
                #endregion

                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    BlockTableRecord acBlkTblRec = (BlockTableRecord)tr.GetObject(dbCurrent.CurrentSpaceId, OpenMode.ForWrite);
                    foreach (ObjectId blockId in blockIds)
                    {
                        BlockReference bRef = (BlockReference)tr.GetObject(blockId, OpenMode.ForRead);
                        Point3d pnt1 = bRef.Position;
                        Curve pl = plines.OrderBy(x => x.GetClosestPointTo(pnt1, false).DistanceTo(pnt1)).First();
                        Point3d pnt2 = pl.GetClosestPointTo(pnt1, false);
                        
                        Point3d midPnt = new Point3d((pnt1.X + pnt2.X) / 2, (pnt1.Y + pnt2.Y) / 2, (pnt1.Z + pnt2.Z) / 2); //середина расстояния
                        Vector3d dimv = new Vector3d(pnt2.X - pnt1.X, pnt2.Y - pnt1.Y, pnt2.Z - pnt1.Z); //вектор расстояния
                        Vector3d normv = dimv.GetPerpendicularVector();  //вектор оффсета                      
                        normv = normv.MultiplyBy(offset / normv.Length); //корректируем длину вектора оффсета на заданную
                        Point3d offsetPnt = midPnt.Add(normv); //точка с оффсетом
                        AlignedDimension drdim = new AlignedDimension(pnt1, pnt2, offsetPnt, "", ObjectId.Null);
                        acBlkTblRec.AppendEntity(drdim);
                        tr.AddNewlyCreatedDBObject(drdim, true);
                    }
                    tr.Commit();
                }
                foreach (Curve pl in plines) pl.Unhighlight();

            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.ToString());
            }
        }

        [CommandMethod("смарткпс")]
        public void cmdSmartLineMatchProperties()
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            //PromptEntityOptions opt = new PromptEntityOptions(;           
            PromptEntityResult res = ed.GetEntity("Выберите исходную полилинию:");
            if (res.Status == PromptStatus.OK)
            {
                ObjectId id = res.ObjectId;
                Database dbCurrent = HostApplicationServices.WorkingDatabase;
                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {
                    PlineObject item;
                    //BlockTableRecord space = (BlockTableRecord)tr.GetObject(dbCurrent.CurrentSpaceId, OpenMode.ForRead);                    
                    Entity entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (entity.GetType() == typeof(Polyline))
                    {
                        DwgObject obj2 = new DwgObject(id);
                        if (obj2.HasExtData && (obj2.ObjectType != "trench"))
                        {
                            item = PlineObject.Open(id);
                        }
                        else return;
                        TypedValue[] values = new TypedValue[] { new TypedValue(0, "LWPOLYLINE") };
                        SelectionFilter filter = new SelectionFilter(values);
                        PromptSelectionOptions opts = new PromptSelectionOptions();

                        //opts.MessageForRemoval = "\nMust be a type of Block!";
                        opts.MessageForAdding = "\nВыберите целевые полилинии: ";
                        //opts.PrepareOptionalDetails = true;
                        //opts.SingleOnly = true;
                        //opts.SinglePickInSpace = true;
                        //opts.AllowDuplicates = false;
                        PromptSelectionResult res2 = default(PromptSelectionResult);
                        res2 = ed.GetSelection(opts, filter);
                        if (res2.Status != PromptStatus.OK)
                            return;
                        try
                        {                           
                            //BlockTable acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                            //BlockTableRecord acBlkTblRec = (BlockTableRecord)tr.GetObject(dbCurrent.CurrentSpaceId, OpenMode.ForWrite);
                            foreach (SelectedObject selPl in res2.Value)
                            {                                
                                Polyline pline = tr.GetObject(selPl.ObjectId, OpenMode.ForWrite, false) as Polyline;

                                item.ObjId = selPl.ObjectId;

                                item.SaveXMLtoCADEntity(item.ToXElement());
                            }                                
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
                        }
                    }

                    tr.Commit();
                }
                
            }
        }

        [CommandMethod("vl_volumes")]
        public void cmdVl_Volumes()
        {
            DBVolumeForm frm = new DBVolumeForm();
            frm.connection = connection;
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(frm);

        }

        private void ImportDwg(string sourceFileName)
        {
            DocumentCollection dm = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            Editor ed = dm.MdiActiveDocument.Editor;
            Database destDb = dm.MdiActiveDocument.Database;
            Database sourceDb = new Database(false, true);            
            try
            {
                /* Загружаем чертеж по ссылке */
                /* Копируем динамический блок во временный каталог*/                
                sourceDb.ReadDwgFile(sourceFileName, System.IO.FileShare.Read, true, "");
                ObjectIdCollection blockIds = new ObjectIdCollection();
                Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = sourceDb.TransactionManager;
                using (Transaction myT = tm.StartTransaction())
                {
                    /* Метка нашей вставки блока */
                    //Handle handle = new Handle(0x215);
                    //ObjectId brefId = ObjectId.Null;
                    //sourceDb.TryGetObjectId(handle, out brefId);
                    //blockIds.Add(brefId);
                    BlockTable sourceDbBlockTable = (BlockTable)myT.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord sourceDbModelSpace = myT.GetObject(sourceDbBlockTable[BlockTableRecord.ModelSpace], 
                        OpenMode.ForRead) as BlockTableRecord;
                    foreach (ObjectId entityId in sourceDbModelSpace)
                    {
                        blockIds.Add(entityId);
                    }                    
                }
                IdMapping mapping = new IdMapping();
                destDb.WblockCloneObjects(blockIds, destDb.CurrentSpaceId,
                mapping, DuplicateRecordCloning.Replace, false);
                //mapping.
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("\nОшибка при копировании: " + ex.Message);
            }
            sourceDb.Dispose();
            //Document _doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            //Editor _ed = _doc.Editor;
            //Database _db = HostApplicationServices.WorkingDatabase;
            //Matrix3d ucs = _ed.CurrentUserCoordinateSystem;
            //string blockname = Path.GetFileNameWithoutExtension(sourceDrawing);
            //try
            //{
            //    using (_doc.LockDocument())
            //    {
            //        using(Database inMemoryDb=new Database(false, true));
            //        {
            //            using (Transaction transaction = _db.TransactionManager.StartTransaction())
            //            {
            //                BlockTable destDbBlockTable = (BlockTable)transaction.GetObject(_db.BlockTableId, OpenMode.ForRead);
            //                BlockTableRecord destDbCurrentSpace = (BlockTableRecord)_db.CurrentSpaceId.GetObject(OpenMode.ForWrite);
                            
            //            }
            //        }
            //    }
            //}
            //catch(System.Exception)
            //{
            //}
            
        }

        private Layout ImportLayoutWithOrWithoutReplace(string fileName, string layoutName)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            if (CheckLayout(layoutName))
            {
                DialogResult res = MessageBox.Show("Лист " + layoutName + " существует. Заменить?",
                    "vl_tools", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    DocumentLock oLock = Autodesk.AutoCAD.ApplicationServices.Application.
                        DocumentManager.MdiActiveDocument.LockDocument();
                    Database dbSource = HostApplicationServices.WorkingDatabase;
                    try
                    {
                        using (Transaction trans = dbSource.TransactionManager.StartTransaction())
                        {
                            DBDictionary dbdLayout = (DBDictionary)trans.GetObject(dbSource.LayoutDictionaryId,
                                OpenMode.ForRead, false, false);

                            foreach (DictionaryEntry deLayout in dbdLayout)
                            {
                                string curLayName = deLayout.Key.ToString().ToLower();
                                if (curLayName == layoutName.ToLower())
                                {
                                    LayoutManager.Current.DeleteLayout(layoutName); // Delete layout.
                                }
                            }
                            trans.Commit();
                        }
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        MessageBox.Show(ex.ToString(), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }
                    catch (System.Exception ex)
                    {

                        MessageBox.Show(ex.ToString(), "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }
                    finally
                    {
                        dbSource.Dispose();
                        oLock.Dispose();
                    }
                }
            }
            string outName = ImportLayout(fileName, layoutName);
            //ed.Regen();   // Updates AutoCAD GUI to relect changes.
            return GetLayout(outName);
        }

        public static bool CheckLayout(string layoutName)
        {
            DocumentLock oLock = Autodesk.AutoCAD.ApplicationServices.Application.
                DocumentManager.MdiActiveDocument.LockDocument();
            Database dbSource = HostApplicationServices.WorkingDatabase;
            try
            {
                using (Transaction trans = dbSource.TransactionManager.StartTransaction())
                {
                    ObjectId idDbdSource = dbSource.LayoutDictionaryId;
                    DBDictionary dbdLayout = (DBDictionary)trans.GetObject(idDbdSource, OpenMode.ForRead, false, false);


                    ObjectIdCollection idc = new ObjectIdCollection();
                    foreach (DictionaryEntry deLayout in dbdLayout)
                    {
                        string curLayName = deLayout.Key.ToString().ToLower();
                        if (curLayName == layoutName.ToLower())
                        {
                            return true;
                        }
                    }

                    trans.Commit();
                }
                return false;

            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.ToString(), "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {

                dbSource.Dispose();
                oLock.Dispose();
            }
        }

        public Layout GetLayout(string layoutName)
        {
            DocumentLock oLock = Autodesk.AutoCAD.ApplicationServices.Application.
                DocumentManager.MdiActiveDocument.LockDocument();
            Database dbSource = HostApplicationServices.WorkingDatabase;
            try
            {
                using (Transaction trans = dbSource.TransactionManager.StartTransaction())
                {
                    ObjectId idDbdSource = dbSource.LayoutDictionaryId;
                    DBDictionary dbdLayout = (DBDictionary)trans.GetObject(idDbdSource, OpenMode.ForRead, false, false);


                    ObjectIdCollection idc = new ObjectIdCollection();
                    foreach (DictionaryEntry deLayout in dbdLayout)
                    {
                        string curLayName = deLayout.Key.ToString().ToLower();
                        if (curLayName == layoutName.ToLower())
                        {
                            ObjectId idLayout = (ObjectId)deLayout.Value;
                            return (Layout)trans.GetObject(idLayout, OpenMode.ForRead);
                        }
                    }

                    trans.Commit();
                }
                return null;

            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.ToString(), "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {

                dbSource.Dispose();
                oLock.Dispose();
            }
        }

        public string ImportLayout(string fileName, string layoutName)
        {
            string retString = "";
            DocumentLock oLock = Autodesk.AutoCAD.ApplicationServices.Application.
                DocumentManager.MdiActiveDocument.LockDocument();
            Database dbSource = new Database(false, false);
            Database db = HostApplicationServices.WorkingDatabase;
            try
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    dbSource.ReadDwgFile(fileName, System.IO.FileShare.Read, true, null);
                    ObjectId idDbdSource = dbSource.LayoutDictionaryId;
                    DBDictionary dbdLayout = (DBDictionary)trans.GetObject(idDbdSource, OpenMode.ForRead, false, false);
                    ObjectId idLayout;

                    ObjectIdCollection idc = new ObjectIdCollection();
                    //Ищем имя листа в исходном файле
                    foreach (DictionaryEntry deLayout in dbdLayout)
                    {
                        string curLayName = deLayout.Key.ToString().ToLower();


                        if (curLayName == layoutName.ToLower())
                        {
                            idLayout = (ObjectId)deLayout.Value;
                            Layout ltr = (Layout)trans.GetObject(idLayout, OpenMode.ForWrite);

                            string copyLayName = layoutName;
                            int i = 1;
                            while (CheckLayout(copyLayName))
                            {
                                copyLayName = layoutName + "_" + i.ToString();
                                ltr.LayoutName = copyLayName;
                                i += 1;
                            }
                            idc.Add(idLayout);
                            retString = copyLayName;
                            break;
                        }
                    }

                    IdMapping im = new IdMapping();
                    db.WblockCloneObjects(idc, db.LayoutDictionaryId, im, DuplicateRecordCloning.Ignore, false);
                    trans.Commit();
                    return retString;
                }

            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return retString;
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.ToString(), "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return retString;
            }
            finally
            {

                dbSource.Dispose();
                oLock.Dispose();
            }
        }

        private ObjectId GetFirstTable(Layout lt)
        {
            ObjectId tblID = ObjectId.Null;
            Database db = HostApplicationServices.WorkingDatabase;
            Transaction trans = db.TransactionManager.StartTransaction();
            try
            {
                BlockTableRecord layoutTableRecord = (BlockTableRecord)lt.
                    Database.TransactionManager.GetObject(lt.BlockTableRecordId, OpenMode.ForWrite, false);
                foreach (ObjectId entityId in layoutTableRecord)
                {
                    object acadObject = lt.Database.TransactionManager.
                        GetObject(entityId, OpenMode.ForRead, false);
                    Table tbl = acadObject as Table;
                    if (tbl != null)
                    {
                        tblID = tbl.ObjectId;
                        break;
                    }
                }
            }
            finally
            {
                trans.Commit();
            }
            return tblID;
        }

        public void ImportBlock(string sourceFileName, string BlockName, string newBlockName="")
        {
            DocumentCollection dm =
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            Editor ed = dm.MdiActiveDocument.Editor;
            Database destDb = dm.MdiActiveDocument.Database;
            Database sourceDb = new Database(false, true);
            try
            {
                // Read the DWG into a side database
                sourceDb.ReadDwgFile(sourceFileName,
                                    System.IO.FileShare.Read, true, "");

                // Create a variable to store the list of block identifiers
                ObjectIdCollection blockIds = new ObjectIdCollection();
                Autodesk.AutoCAD.DatabaseServices.TransactionManager tm =
                  sourceDb.TransactionManager;

                using (Transaction myT = tm.StartTransaction())
                {
                    // Open the block table
                    BlockTable bt = (BlockTable)tm.GetObject(sourceDb.BlockTableId,
                                                OpenMode.ForRead, false);
                    if (!bt.Has(BlockName))
                    {
                        throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.InvalidBlockName,
                            "\nBlock \"" + BlockName + "\" not found in \"" + sourceFileName + "\"\n");
                    }
                    BlockTableRecord btr = (BlockTableRecord)tm.GetObject(bt[BlockName],
                        OpenMode.ForWrite, false);
                    if(newBlockName!="") btr.Name = newBlockName;                 
                    if (!btr.IsAnonymous && !btr.IsLayout) blockIds.Add(btr.ObjectId);
                    btr.Dispose();
                }

                // Copy blocks from source to destination database

                IdMapping mapping = new IdMapping();

                sourceDb.WblockCloneObjects(blockIds, destDb.BlockTableId, mapping,
                    DuplicateRecordCloning.Ignore, false);
            }

            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                throw ex;
            }
            sourceDb.Dispose();
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

        /// <summary>
        /// Определяет, лежит ли точка в заданном допуском расстоянии от другой точки
        /// </summary>
        /// <param name="pnt1">точка, от которой рассчитывается расстояние</param>
        /// <param name="pnt2">точка, до которой рассчитывается расстояние</param>
        /// <param name="tolerance">допуск</param>
        /// <returns>лежит ил нет</returns>
        private bool IsPointNear(Point3d pnt1, Point3d pnt2, double tolerance)
        {
            try
            {
                return (pnt1 - pnt2).Length <= tolerance;
            }
            catch { }            
            // Otherwise we return false
            return false;            
        }

        /// <summary>
        /// checks the position of point
        /// </summary>
        /// <param name="cv"></param>
        /// <param name="pt"></param>
        /// <returns></returns>
        private bool IsPointOnCurveGCP(Curve cv, Point3d pt)
        {
            try
            {
                // Return true if operation succeeds
                Point3d p = cv.GetClosestPointTo(pt, false);
                return (p - pt).Length <= Tolerance.Global.EqualPoint;
            }
            catch { }
            // Otherwise we return false
            return false;
        }

        //private double GetDistFromPointToPolyline(Polyline pline, Point2d pnt)
        //{            
        //    List<double> distList = new List<double>();

        //    for (int i = 0; i < pline.NumberOfVertices - 1; i++)
        //    {
        //        Point2d curPnt = pline.GetPoint2dAt(i);
        //        Point2d nextPnt = pline.GetPoint2dAt(i + 1);
        //        Vector2d curVector = new Vector2d(nextPnt.X - curPnt.X, nextPnt.Y - curPnt.Y);
        //        pline.get


        //    }

        //}
    }

    //class AttInfo
    //{
    //    private Point3d _pos;
    //    private Point3d _aln;
    //    private bool _aligned;

    //    public AttInfo(Point3d pos, Point3d aln, bool aligned)
    //    {
    //        _pos = pos;
    //        _aln = aln;
    //        _aligned = aligned;
    //    }

    //    public Point3d Position
    //    {
    //        set { _pos = value; }
    //        get { return _pos; }
    //    }

    //    public Point3d Alignment
    //    {
    //        set { _aln = value; }
    //        get { return _aln; }
    //    }

    //    public bool IsAligned
    //    {
    //        set { _aligned = value; }
    //        get { return _aligned; }
    //    }
    //}
}
