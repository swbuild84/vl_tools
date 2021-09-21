using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace vl_tools
{
    public class VLRoadLightClass
    {
        private Polyline _plRight;
        private Polyline _plLeft;
        private Polyline _plAxis;
        private double step;
        private double tolerance = 0.1;


        public static Point3d GetPlPointAtDist(Polyline pline, Point3d startPnt, double step, double tolerance)
        {
            try
            {
                Point3d OnPLPoint = pline.GetClosestPointTo(startPnt, false);
                double dist = pline.GetDistAtPoint(OnPLPoint);
                Point3d CurPoint = OnPLPoint;
                dist += step;
                if (dist <= pline.Length)
                {
                    Point3d NextPnt = pline.GetPointAtDist(dist);
                    double calcStep = (NextPnt - CurPoint).Length;
                    while (Math.Abs(calcStep - step) > tolerance)
                    {
                        double delta = step - calcStep;
                        dist += delta / 2;
                        NextPnt = pline.GetPointAtDist(dist);
                        calcStep = (NextPnt - CurPoint).Length;
                    }
                    return NextPnt;
                }
                else throw new VLRoadLightInvalidInputException("Дистанция за пределами полилинии");

            }
            catch (Exception)
            {
                throw;
            }

        }

    }

    public class VLRoadLightInvalidInputException : Exception
    {
        public VLRoadLightInvalidInputException(string message) : base(message) { }        
    }
}
