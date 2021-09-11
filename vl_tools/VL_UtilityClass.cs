using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace vl_tools
{
    public static class VL_UtilityClass
    {

        public static void ImportDwg(string sourceFileName)
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

        public static Layout ImportLayoutWithOrWithoutReplace(string fileName, string layoutName)
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

        public static Layout GetLayout(string layoutName)
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

        public static string ImportLayout(string fileName, string layoutName)
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

        public static ObjectId GetFirstTable(Layout lt)
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

        public static void ImportBlock(string sourceFileName, string BlockName, string newBlockName = "")
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
                    if (newBlockName != "") btr.Name = newBlockName;
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
        public static bool IsPointNear(Point3d pnt1, Point3d pnt2, double tolerance)
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
        public static bool IsPointOnCurveGCP(Curve cv, Point3d pt)
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

        public static Polyline GetPolyline(string promptString)
        {
            try
            {
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Database dbCurrent = HostApplicationServices.WorkingDatabase;
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                PromptEntityOptions opt = new PromptEntityOptions("\n" + promptString + ":");
                opt.SetRejectMessage("\nВыбранный объект - не полилиния!");
                opt.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                PromptEntityResult res = ed.GetEntity(opt);
                if (res.Status == PromptStatus.OK)
                {
                    ObjectId id = res.ObjectId;
                    using (DocumentLock acLckDoc = doc.LockDocument())
                    {
                        using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                        {
                            Entity entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (entity.GetType() == typeof(Polyline))
                            {
                                return (Polyline)entity;
                            }
                            else throw new System.Exception("Выбранный объект - не полилиния!");
                        }
                    }
                }
                else
                {
                    throw new System.Exception("Выбор отменен пользователем");
                }

            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.ToString());
                return null;
            }
        }

    }
}
