using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace vl_tools
{
    /// <summary>
    /// Класс для записи настроек трассы в файл чертежа
    /// </summary>
    public class VLFileOptions
    {
        /// <summary>
        /// Имя словаря
        /// </summary>
        private const string nodName = "vl_toolsNOD";
        /// <summary>
        /// Флаг монитора трассы
        /// </summary>
        public bool VL_PROFILE_MONITOR;
        /// <summary>
        /// Хэндл полилинии трассы
        /// </summary>
        public Handle profileTraceHandle;
        /// <summary>
        /// Начальный пикет
        /// </summary>
        public double profileStartPicket;
        /// <summary>
        /// Начальный километр трассы
        /// </summary>
        public double profileStartKilometer;

        /// <summary>
        /// Запись настроек программы в текущий файл чертежа
        /// </summary>
        public void Write()
        {
            try
            {
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Database dbCurrent = HostApplicationServices.WorkingDatabase;
                using (Transaction tr = dbCurrent.TransactionManager.StartTransaction())
                {                    
                    DBDictionary nod = (DBDictionary)tr.GetObject(dbCurrent.NamedObjectsDictionaryId, OpenMode.ForWrite);
                    if (nod.Contains(nodName))
                    {
                        ObjectId myDataId = nod.GetAt(nodName);
                        Xrecord xRec = (Xrecord)tr.GetObject(myDataId, OpenMode.ForWrite);
                        xRec.Data = new ResultBuffer(new TypedValue((int)DxfCode.Bool, VL_PROFILE_MONITOR),
                                                          new TypedValue((int)DxfCode.Handle, profileTraceHandle),
                                                          new TypedValue((int)DxfCode.Int32, profileStartPicket),
                                                          new TypedValue((int)DxfCode.Int32, profileStartKilometer));
                    }
                    else
                    {
                        Xrecord myXrecord = new Xrecord();
                        myXrecord.Data = new ResultBuffer(new TypedValue((int)DxfCode.Bool, VL_PROFILE_MONITOR),
                                                          new TypedValue((int)DxfCode.Handle, profileTraceHandle),
                                                          new TypedValue((int)DxfCode.Int32, profileStartPicket),
                                                          new TypedValue((int)DxfCode.Int32, profileStartKilometer));
                        nod.SetAt(nodName, myXrecord);
                        tr.AddNewlyCreatedDBObject(myXrecord, true);
                    }                    
                    tr.Commit();
                }
            }
            catch (System.Exception)
            {
                //Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.ToString());                
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool TryToRead()
        {
            try
            {
                Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
                {
                    // Find the NOD in the database
                    DBDictionary nod = (DBDictionary)trans.GetObject(acCurDb.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (!nod.Contains(nodName)) return false; //словарь не найден
                    ObjectId myDataId = nod.GetAt(nodName);
                    Xrecord readBack = (Xrecord)trans.GetObject(myDataId, OpenMode.ForRead);
                    TypedValue[] options = readBack.Data.AsArray();
                    VL_PROFILE_MONITOR = System.Convert.ToBoolean(options[0].Value);
                    profileTraceHandle= new Handle(System.Convert.ToInt64(options[1].Value.ToString(), 16));
                    profileStartPicket = System.Convert.ToDouble(options[2].Value);
                    profileStartKilometer = System.Convert.ToDouble(options[3].Value);
                    trans.Commit();
                } // using
                return true;
            }
            catch (System.Exception)
            {
               // Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.ToString());
               
                return false;
            }
        }
    

    }
}
