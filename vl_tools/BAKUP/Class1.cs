﻿using System;
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

[assembly: CommandClass(typeof(vl_tools.Commands))]
namespace vl_tools
{
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

        public void Initialize()
        {
            DemandLoading.RegistryUpdate.RegisterForDemandLoading();
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage(this.GetType().Module.Name + " загружен и добавлен в автозагрузку\n");
            ed.WriteMessage("Для удаления из автозагрузки используйте команду vl_tools_REMAUTO\n");

            #region CONFIG
            try
            {
                // location update
                Assembly assem = Assembly.GetExecutingAssembly();
                string name = assem.ManifestModule.Name;
                string appPath = assem.Location.Replace(name, "");
                _TemplatePath = appPath;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            #endregion
        }

        public void Terminate()
        {
            //throw new NotImplementedException();
        }

        #endregion

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
                if ((s[i] > 47 && s[i] < 58) || (s[i] == 46))
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

        [CommandMethod("provis")]
        public void cmdProvis()
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTableRecord acBlkTblRec = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);                    
                }
            }

            catch (System.Exception ex)
            {
                ed.WriteMessage("Error: " + ex.Message + "\n" + ex.StackTrace);
            }
        }

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
                List<string> array = new List<string>();
                
                string[] fields = lines[0].Split('\t');
                
                foreach (string s in lines)
                {
                    string[] cols;
                    cols = s.Split('\t');
                    
                    array.AddRange(cols);
                }
                array.Remove("");
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
                string[] skwrds = new string[] { "1", "2", "3" };
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
                            if (fld == "coord_X")
                            {
                                Xcoord = GetDoubleFString(value);
                            }
                            if (fld == "coord_Y")
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

        static Point2d PolarPoints(Point2d pPt, double dAng, double dDist)
        {
            return new Point2d(pPt.X + dDist * Math.Cos(dAng),
                               pPt.Y + dDist * Math.Sin(dAng));
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
            opts.MessageForAdding = "\nSelect a polyline: ";
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
                        Polyline pline = tr.GetObject(selPl.ObjectId, OpenMode.ForRead, false) as Polyline;
                        Point2d prevPnt = pline.GetPoint2dAt(0);
                        Point2d curPnt = new Point2d();

                        for (int i = 1; i < pline.NumberOfVertices; i++)
                        {
                            curPnt = pline.GetPoint2dAt(i);
                            double dist = prevPnt.GetDistanceTo(curPnt);
                            int ncols = Convert.ToInt32(Math.Ceiling(dist / step));
                            double resStep = dist / ncols;
                            double dAng = prevPnt.GetVectorTo(curPnt).Angle;
                            //Промежуточные точки
                            for (int j = 0; j < ncols-1; j++)
                            {
                                Point2d newPnt = PolarPoints(prevPnt, dAng, resStep);
                                DBPoint acPoint = new DBPoint(new Point3d(newPnt.X, newPnt.Y, 0));
                                acPoint.SetDatabaseDefaults();
                                space.AppendEntity(acPoint);
                                tr.AddNewlyCreatedDBObject(acPoint,true);
                                prevPnt = newPnt;
                            }
                            prevPnt = curPnt;

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
                    
                int cntr = 1;
                foreach (IGrouping<string, Abonent> g in TPAbGroups)
                {
                    //Show dlg
                    TPForm frm = new TPForm();
                    frm.m_group = g;
                    frm.templatePath = _TemplatePath + "ТП";
                    DialogResult TPres = Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(frm);
                    if (TPres != DialogResult.OK) continue;

                    //Получаем ТП и автоматы
                    DataGridView grid = frm.dataGridView1;
                    DataGridView grid2 = frm.dataGridView2;
                    List<int> list2 = frm.dgv1Indexes;

                    List<string> LAutomates = new List<string>();
                    LAutomates = frm.automates;
                    string filePath = frm.SelectedFilePath;
                    string blockName = frm.SelectedBlockName;
                    bool flag15 = frm.checkBoxIsExistingTP.Checked;

                    // Create the value GUIDs.
                    string newBlockName = Guid.NewGuid().ToString();
                    ImportBlock(filePath, blockName, newBlockName);
                    BlockReference br = BlockJig(newBlockName, new Hashtable());
                    //br.Explode(new DBObjectCollection());
                    
                    Transaction trans = dbCurrent.TransactionManager.StartTransaction();
                    {
                        try
                        {
                            BlockTable bt = (BlockTable)trans.GetObject(dbCurrent.BlockTableId,
                                      OpenMode.ForRead, false);
                            BlockTableRecord record = (BlockTableRecord)trans.GetObject(br.BlockTableRecord,
                                OpenMode.ForWrite, false);
                            Table TblFiders = null;
                            Table TblElems = null;
                            //Опросник
                            Table TblOprosnik = null;
                            Table TblTP = null;

                            foreach (ObjectId entityId in record)
                            {
                                DBObject acadObject = trans.GetObject(entityId, OpenMode.ForRead, false);
                                BlockReference reference4 = acadObject as BlockReference;
                                if (reference4 != null)
                                {
                                    AttributeCollection attributes2 = reference4.AttributeCollection;
                                    bool flag18 = false;
                                    foreach (ObjectId id4 in attributes2)
                                    {
                                        AttributeReference reference5 = (AttributeReference)trans.GetObject(id4, OpenMode.ForWrite);
                                        string str8 = "";
                                        if (grid2.Rows[0].Cells[1].Value != null)
                                        {
                                            str8 = grid2.Rows[0].Cells[1].Value.ToString();
                                        }
                                        if (reference5.Tag == "НАЗВАНИЕ_ЧЕРТЕЖА")
                                        {
                                            reference5.TextString = "Схема однолинейная принципиальная сети 0,4 кВ (" + str8 + ")";
                                            flag18 = true;
                                        }
                                    }
                                    foreach (ObjectId id5 in attributes2)
                                    {
                                        AttributeReference reference6 = (AttributeReference)trans.GetObject(id5, OpenMode.ForWrite);
                                        if (reference6.Tag == "ЛИСТ")
                                        {
                                            if (flag18)
                                            {
                                                if (flag15)
                                                {
                                                    reference6.TextString = num.ToString();
                                                }
                                                else
                                                {
                                                    reference6.TextString = num.ToString() + ".1";
                                                }
                                            }
                                            else
                                            {
                                                reference6.TextString = num.ToString() + ".2";
                                            }
                                        }
                                    }
                                }

                                Table ACADtbl = acadObject as Table;
                                if (ACADtbl != null)
                                {
                                    //check what it is?
                                    if (ACADtbl.Cells[0, 1].TextString == "Номер группы, линии (фидер)")
                                    {
                                        TblFiders = trans.GetObject(entityId, OpenMode.ForWrite, false) as Table;
                                    }
                                    if (ACADtbl.Cells[0, 0].TextString == "Ведомость силового оборудования")
                                    {
                                        TblElems = trans.GetObject(entityId, OpenMode.ForWrite, false) as Table;
                                    }
                                    if (ACADtbl.Cells[0, 0].TextString == "Номер ТП")
                                    {
                                        TblTP = trans.GetObject(entityId, OpenMode.ForWrite, false) as Table;
                                    }
                                    //Находим опросник
                                    if (ACADtbl.Cells[0, 0].TextString.Contains("ОПРОСНЫЙ ЛИСТ"))
                                    {
                                        TblOprosnik = trans.GetObject(entityId, OpenMode.ForWrite, false) as Table;
                                    }

                                }
                            }
                            //Fill tables
                            if (TblFiders != null)
                            {
                                TblFiders.InsertColumns(2, 25.0, LAutomates.Count);
                                int col = 2;

                                foreach (int index in list2)
                                //for (int i = 0; i < LAutomates.Count; i++)
                                {
                                    if (grid.Rows[9].Cells[index].Value != null)
                                        TblFiders.Cells[0, col].TextString = grid.Columns[index].Name;
                                    if (grid.Rows[9].Cells[index].Value != null)
                                        TblFiders.Cells[1, col].TextString = grid.Rows[9].Cells[index].Value.ToString();
                                    if (grid.Rows[8].Cells[index].Value != null)
                                        TblFiders.Cells[2, col].TextString = grid.Rows[8].Cells[index].Value.ToString() + " А";
                                    if (grid.Rows[6].Cells[index].Value != null)
                                        TblFiders.Cells[3, col].TextString = grid.Rows[6].Cells[index].Value.ToString();
                                    if (grid.Rows[7].Cells[index].Value != null)
                                        TblFiders.Cells[4, col].TextString = grid.Rows[7].Cells[index].Value.ToString();
                                    if (grid.Rows[5].Cells[index].Value != null)
                                        TblFiders.Cells[5, col].TextString = grid.Rows[5].Cells[index].Value.ToString();
                                    if (grid.Rows[1].Cells[index].Value != null)
                                        TblFiders.Cells[6, col].TextString = grid.Rows[1].Cells[index].Value.ToString();
                                    if (grid.Rows[0].Cells[index].Value != null)
                                        TblFiders.Cells[7, col].TextString = grid.Rows[0].Cells[index].Value.ToString() + " кВт";
                                    if (grid.Rows[3].Cells[index].Value != null)
                                        TblFiders.Cells[8, col].TextString = grid.Rows[3].Cells[index].Value.ToString();
                                    if (grid.Rows[4].Cells[index].Value != null)
                                        TblFiders.Cells[9, col].TextString = grid.Rows[4].Cells[index].Value.ToString() + " кВт";
                                    if (grid.Rows[2].Cells[index].Value != null)
                                        TblFiders.Cells[10, col].TextString = grid.Rows[2].Cells[index].Value.ToString();
                                    col++;
                                }
                            }
                            if (TblElems != null)
                            {
                                int count = TblElems.Rows.Count;
                                double height = TblElems.Rows[count - 1].Height;
                                int i = 1;
                                int row = 0;
                                foreach (int index in list2)
                                {
                                    if (!((grid.Rows[4].Cells[index].Value == null) & flag15))
                                    {
                                        TblElems.InsertRows((count - 1) + row, height, 1);
                                        TblElems.Cells[(count - 1) + row, 0].TextString = "QF" + i.ToString();
                                        string marka = "ВА51-35 ";
                                        if (Convert.ToDouble(LAutomates[i - 1]) > 250.0)
                                        {
                                            marka = "ВА57-39 ";
                                        }
                                        TblElems.Cells[(count - 1) + row, 1].TextString = marka + LAutomates[i - 1] + " А";
                                        TblElems.Cells[(count - 1) + row, 1].Alignment = CellAlignment.MiddleLeft;
                                        TblElems.Cells[(count - 1) + row, 2].TextString = "1";
                                        row++;
                                    }
                                    i++;
                                }

                            }
                            if (TblTP != null)
                            {
                                if (grid2.Rows[0].Cells[1].Value != null)
                                {
                                    TblTP.Cells[0, 1].TextString = grid2.Rows[0].Cells[1].Value.ToString();
                                }
                                if (grid2.Rows[1].Cells[1].Value != null)
                                {
                                    TblTP.Cells[1, 1].TextString = grid2.Rows[1].Cells[1].Value.ToString();
                                }
                                if (grid2.Rows[3].Cells[1].Value != null)
                                {
                                    TblTP.Cells[2, 1].TextString = grid2.Rows[3].Cells[1].Value.ToString();
                                }
                                double Tok = (Convert.ToDouble(grid2.Rows[5].Cells[1].Value) / 1.73) / 0.38;
                                TblTP.Cells[3, 1].TextString = string.Format("{0:f2}", Tok) + " А";
                            }

                            //Заполняем опросник
                            if (TblOprosnik != null)
                            {
                                TblOprosnik.Cells[0, 0].TextString = "ОПРОСНЫЙ ЛИСТ №" + cntr.ToString() +
                                    " (" + grid2.Rows[0].Cells[1].Value.ToString() + ")" + "\r\n"
                                    +  "на трансформаторную подстанцию наружной установки";
                                List<string> nominals = new List<string>();
                                for (int i = 2; i < TblOprosnik.Columns.Count; i++)
                                {
                                    string cellTxt=TblOprosnik.Cells[24, i].TextString;
                                    if (cellTxt != "") nominals.Add(cellTxt);
                                }

                                foreach (string autm in LAutomates) if (!nominals.Contains(autm)) 
                                    MessageBox.Show("Автомат " + autm + " не встречается в опроснике!");

                                for (int i = 0; i < nominals.Count(); i++)
                                {
                                    var slctd = from s in LAutomates.AsEnumerable()
                                                where s.Equals(nominals[i])
                                                select s;
                                    int scnt = slctd.Count();
                                    if (scnt > 0) TblOprosnik.Cells[25, i + 2].TextString = scnt.ToString();
                                }                                
                            }
                        }
                        finally
                        {
                            trans.Commit();
                        }
                        cntr++;
                        num++;  
                    }                   
                }
                ed.Regen();  
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }            
        }


        [CommandMethod("ESP_CHECK")]
        public void CheckAbonentsSQL()
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

        [CommandMethod("SL_gredit")]
        public void SmartLineGroupEdit()
        {
            try
            {
                List<BlockObject> selDwgObjects = new List<BlockObject>();  //список выбранных объектов

                //string appName = "ESMT_LEP_v1.0";
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                Database db = doc.Database;               

                var tv = new TypedValue(0, "INSERT");
                var sf = new SelectionFilter(new TypedValue[] { tv });

                // Ask the user to select (filtered) entities

                var res = ed.GetSelection(sf);

                if (res.Status != PromptStatus.OK)
                    return;

                using (var tr = doc.TransactionManager.StartTransaction())
                {
                    try
                    {
                        ObjectId[] ids = res.Value.GetObjectIds();
                        
                        foreach (ObjectId id in ids)
                        {
                            BlockObject dwgObj = new BlockObject(id);
                            if (!dwgObj.HasExtData) continue;
                            selDwgObjects.Add(dwgObj);

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
            //Вставляет блоки в вершины полилинии
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;
            TypedValue[] values = new TypedValue[] { new TypedValue(0, "LWPOLYLINE") };
            SelectionFilter filter = new SelectionFilter(values);

            PromptSelectionOptions opts = new PromptSelectionOptions();
            opts.AllowDuplicates = false;

            opts.Keywords.Add("Options");            

            string kws = opts.Keywords.GetDisplayString(true);
            opts.MessageForAdding = "\nВыберите трассу ВЛ или " + kws;
            opts.KeywordInput +=
               delegate(object sender, SelectionTextInputEventArgs e)
               {
                   ed.WriteMessage("\nKeyword entered: {0}", e.Input);
               };
            PromptSelectionResult res = ed.GetSelection(opts, filter);
            if (res.Status == PromptStatus.OK)
            {
                ed.WriteMessage(res.ToString());
            }            
            else
            {
                ed.WriteMessage("Прервано пользователем\n");
                return;
            }

            DataSet ds = new DataSet();
            try
            {
                //Init database frame                
                System.Data.DataTable table = new System.Data.DataTable();
                table.Columns.Add(new System.Data.DataColumn("GROUP_NAME", typeof(string), "", MappingType.Attribute));
                table.Columns.Add(new System.Data.DataColumn("ANGLE", typeof(string), "", MappingType.Attribute));
                table.Columns.Add(new System.Data.DataColumn("DWGNAME", typeof(string), "", MappingType.Attribute));
                table.Columns.Add(new System.Data.DataColumn("BLOCKNAME", typeof(string), "", MappingType.Attribute));
                ds.Tables.Add(table);
                //table.Rows.Add("0,4 кВ", 1, "ВЛ.dwg", "04_Пром");
                string DBPath = _TemplatePath + "ВЛ\\ВЛ.xml";
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
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
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
                            var rows=ds.Tables[0].AsEnumerable();
                            var BiggestAngles = rows.Where(row => Convert.ToDouble(row.Field<string>("ANGLE")) >= dAng);
                            if (BiggestAngles.Count() > 0)
                            {
                                var arow = BiggestAngles.OrderBy(r => Convert.ToDouble(r.Field<string>("ANGLE"))).First();
                                if (arow != null)
                                {
                                    string filePath = _TemplatePath + "ВЛ\\" + arow.Field<string>("DWGNAME");
                                    string blockName = arow.Field<string>("BLOCKNAME");
                                    //Create the value GUIDs.                           
                                    string newBlockName = Guid.NewGuid().ToString();
                                    ImportBlock(filePath, blockName, newBlockName);
                                    Point2d pnt1 = pline.GetPoint2dAt(i);
                                    BlockTable acBlkTbl = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acBlkTbl[newBlockName], OpenMode.ForRead);
                                    BlockReference br = new BlockReference(new Point3d(pnt1.X, pnt1.Y, 0), btr.ObjectId);
                                    space.AppendEntity(br);
                                    tr.AddNewlyCreatedDBObject(br, true);                                    
                                }
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



        [CommandMethod("смарткпс")]
        public void SmartLineMatchProperties()
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            //PromptEntityOptions opt = new PromptEntityOptions(;           
            PromptEntityResult res = ed.GetEntity("Select polyline:");
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
                        opts.MessageForAdding = "\nSelect a polyline: ";
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
