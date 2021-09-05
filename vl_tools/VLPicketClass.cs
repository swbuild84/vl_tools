using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace vl_tools
{
    enum ESide
    {
        Left, Middle, Right
    }
    
    class VLPicketClass:strigVariables
    {
        public VLFileOptions _FileOpt;        
        private Polyline _pl;
        private double _picket;
        private double _kilometer;
        private double _offset;
        public string _sPicket;
        public string _sKilometer;
        private string _sDist;
        public string _sOffset;
        
        public ESide _side;
        public string _sSide;
        public string Dist => _sDist;
        
        //private const string _pickeVarName = "[ПК]";
        //private const string _kilometerVarName = "[КМ]";
        //private const string _offsetVarName = "[СТОРОНА_СМЕЩЕНИЯ]";
        //private const string _offsetVarName = "[ДИСТАНЦИЯ_СМЕЩЕНИЯ]";

        public VLPicketClass(VLFileOptions FileOpt)
        {
            //_sPicket = "";
            //_sKilometer = "";
            //_sOffset = "";
            _FileOpt = FileOpt;
            Init();    
            
        }

        private void Init()
        {
            try
            {
                Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                Handle hndl = _FileOpt.profileTraceHandle;
                using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
                {
                    ObjectId id = new ObjectId();
                    if (!acCurDb.TryGetObjectId(hndl, out id)) throw new Exception();
                    Polyline pline = trans.GetObject(id, OpenMode.ForRead, false) as Polyline;
                    if (pline == null) return;
                    _pl = pline.Clone() as Polyline;
                    _pl.Elevation = 0;
                    trans.Commit();
                } // using  
            }
            catch (Exception ex)
            {               
            }   
        }

        public override string ToString()
        {

            return "\n" + _sPicket + "\n" + _sKilometer + "\n"+_sSide + " " + _sOffset + "\n";
        }

        /// <summary>
        /// Возвращает косое произведение векторов, позволяющее определить, с какой стороны лежит точка от вектора
        /// </summary>
        /// <param name="Pnt1">первая точка прямой</param>
        /// <param name="Pnt2">вторая точка прямой</param>
        /// <param name="Pnt3">точка для которой определяется сторона относительно прямой</param>
        /// <returns>0-точка на прямой, меньше 0 - точка слева, больше 0 - точка справа</returns>
        public static double PointToPLineSide(Point2d Pnt1, Point2d Pnt2, Point2d Pnt3)
        {
            //D = (х3 - х1) * (у2 - у1) - (у3 - у1) * (х2 - х1)
            return (Pnt3.X - Pnt1.X) * (Pnt2.Y - Pnt1.Y) - (Pnt3.Y - Pnt1.Y) * (Pnt2.X - Pnt1.X);
        }

        public void Calc(Point3d pnt, int picket_round, int offset_round)
        {
            try
            {
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;                
                Point3d pnt2D = new Point3d(pnt.X, pnt.Y, 0);
                Point3d onPlPnt = _pl.GetClosestPointTo(pnt2D, true);
                //double param = _pl.GetParameterAtPoint(onPlPnt);
                Vector3d dist = onPlPnt - pnt;
                _picket = _pl.GetDistAtPoint(onPlPnt) + _FileOpt.profileStartPicket;
                _kilometer = _pl.GetDistAtPoint(onPlPnt) + _FileOpt.profileStartKilometer;
                _offset= onPlPnt.DistanceTo(pnt2D);
                

                var numPicket = System.Math.Truncate(_picket / 100);                
                var addPicket = _picket % 100;
                
                double meters = Math.Truncate(addPicket);   //целая часть                
                string dec = String.Format("{0:f" + picket_round.ToString() + "}", addPicket);
                int pad = 0;
                if (picket_round == 0)
                {
                    pad = 2;
                }
                else
                {
                    pad = picket_round + 3;
                }
                dec = dec.PadLeft(pad, '0');
                _sPicket = "ПК" + numPicket.ToString() + "+" + dec;
                //_sPicket = $"ПК{numPicket}+{meters:00}.{dec}";
                
                var numKM = System.Math.Truncate(_kilometer / 1000);
                var addKM = _kilometer % 1000;
                string dec2 = String.Format("{0:f" + picket_round.ToString() + "}", addKM);
                dec2 = dec2.PadLeft(pad+1, '0');
                _sKilometer = "КМ" + numKM.ToString() + "+" + dec2;

                //double meters2= Math.Truncate(addKM);
                //string dec2 = String.Format("{0:f" + picket_round.ToString() + "}", addKM - meters2);
                //dec2 = dec2.Substring(2);
                //_sKilometer= $"КМ{numKM}+{meters2:000}.{dec2}";

                //_sKilometer = String.Format("КМ{0}+{1:f" + picket_round.ToString() + "}", numKM, addKM);
                _sDist = String.Format("{0:f" + picket_round.ToString() + "}", _picket);
                _sOffset = String.Format("{0:f" + offset_round.ToString() + "}", _offset);
                //if (_offset == 0)
                //{
                //    _sOffset = "";
                //}
                //else
                //{
                //    _sOffset = String.Format("{0:f" + offset_round.ToString() + "}", _offset);
                //}


                double distPK = _pl.GetDistAtPoint(onPlPnt);
                Point3d prevPnt;
                Point3d nextPnt;
                if (distPK>=_pl.Length-0.001)
                {
                    prevPnt = _pl.GetPointAtDist(distPK - 0.001);
                    nextPnt = onPlPnt;
                }
                else
                {
                    prevPnt = onPlPnt;
                    nextPnt = _pl.GetPointAtDist(distPK + 0.001);
                }               
                double side = PointToPLineSide(new Point2d(prevPnt.X, prevPnt.Y), new Point2d(nextPnt.X, nextPnt.Y), new Point2d(pnt2D.X, pnt2D.Y));
                if (side == 0) _side=ESide.Middle;
                if (side < 0) _side=ESide.Left;
                if (side > 0) _side=ESide.Right;
                if (side == 0) _sSide = "";
                if (side < 0) _sSide = "Слева";
                if (side > 0) _sSide = "Справа";
                
                //double param = _pl.GetParameterAtPoint(onPlPnt);
            }
            catch (Exception ex)
            {
                
            }

        }

        public string Interpret(string expression)
        {
            string res = expression.Replace(_pickeVarName, _sPicket);
            return res;
        }
    }
}
