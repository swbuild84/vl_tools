using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace vl_tools
{
    public partial class DrawCatenaryForm : Form
    {
        private const double _g = 9.80665;
        private double _DivisionGammaSigma = 0;
        private double _fm = 0; //провис
        private double _hscale;
        private double _vscale;
        private double _hgab;       

        public DrawCatenaryForm()
        {
            InitializeComponent();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            GroupBoxesDisabled();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            GroupBoxesDisabled();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            GroupBoxesDisabled();
        }

        private void GroupBoxesDisabled()
        {
            groupBox1.Enabled = radioButton1.Checked;
            groupBox2.Enabled = radioButton2.Checked;
            groupBox3.Enabled = radioButton3.Checked;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            try
            {
                if (!SucsessCheckInput()) return;

                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Database db = HostApplicationServices.WorkingDatabase;
                this.Hide();
                PromptPointOptions prPntOpt = new PromptPointOptions("\nУкажите  первую точку: ");
                PromptPointResult prPntRes = ed.GetPoint(prPntOpt);
                if (prPntRes.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); this.DialogResult = DialogResult.OK; return; }
                Point3d Pnt1 = prPntRes.Value;

                int nSpan = 0; //четчик пролетов
                while (true)
                {
                    //Get next point
                    prPntOpt = new PromptPointOptions("\nУкажите  следующую точку: ");
                    prPntRes = ed.GetPoint(prPntOpt);
                    if (prPntRes.Status != PromptStatus.OK) { ed.WriteMessage("Programm was cancelled"); this.DialogResult = DialogResult.OK; return; }
                    Point3d Pnt2 = prPntRes.Value;
                    
                    //Calc vertices
                    double L = (Pnt2.X - Pnt1.X) * _hscale / 1000;
                    double deltaH = (Pnt1.Y - Pnt2.Y) * _vscale / 1000;

                    //счетчик пролетов
                    nSpan++;
                    if (_fm > 0 && nSpan == 1) _DivisionGammaSigma = 8 * _fm / (L * L); //задан провис, нужно найти исходные для механики!                    

                    const double npoints = 30;
                    double step = L / npoints;

                    Point3dCollection points = new Point3dCollection();
                    Point3dCollection pointsOffset = new Point3dCollection();

                    for (double x = 0; x < L; x += step)
                    {
                        double fx = _DivisionGammaSigma * x / 2 * (L - x) + x * deltaH / L;
                        double PntX = Pnt1.X + x * 1000 / _hscale;
                        double PntY = Pnt1.Y - fx * 1000 / _vscale;
                        Point3d pnt = new Point3d(PntX, PntY, 0).TransformBy(ed.CurrentUserCoordinateSystem);
                        Point3d pntOffset = new Point3d(PntX, PntY - _hgab * 1000 / _vscale, 0).TransformBy(ed.CurrentUserCoordinateSystem);
                        points.Add(pnt);
                        pointsOffset.Add(pntOffset);
                    }
                    //add last points
                    points.Add(Pnt2.TransformBy(ed.CurrentUserCoordinateSystem));
                    Point3d Offset = new Point3d(Pnt2.X, Pnt2.Y - _hgab * 1000 / _vscale, 0).TransformBy(ed.CurrentUserCoordinateSystem);
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private bool SucsessCheckInput()
        {
            if(groupBox1.Enabled)
            {
                try
                {
                    _fm = Convert.ToDouble(textBoxFm.Text);                    
                }
                catch (FormatException)
                {
                    MessageBox.Show("Неверный формат!");
                    return false;
                }
            }
            if (groupBox2.Enabled)
            {
                try
                {
                    double koef1 = 0;
                    switch(comboBoxTensionUnit.SelectedIndex)
                    {
                        case 0:
                            koef1 = 1;
                            break;
                        case 1:
                            koef1 = _g;
                            break;
                        case 2:
                            koef1 = 10;
                            break;
                        case 3:
                            koef1 = 1000*_g;
                            break;
                        default:
                            MessageBox.Show("Выберите единицы измерения!");
                            return false;                            
                    }
                    double tension = koef1 * Convert.ToDouble(textBoxTension.Text);

                    double koef2 = 0;
                    switch (comboBoxLoadUnit.SelectedIndex)
                    {
                        case 0:
                            koef2 = 1;
                            break;
                        case 1:
                            koef2 = _g;
                            break;
                        case 2:
                            koef2 = 10;
                            break;                        
                        default:
                            MessageBox.Show("Выберите единицы измерения!");
                            return false;                            
                    }
                    double load = koef2 * Convert.ToDouble(textBoxLoad.Text);
                    _DivisionGammaSigma = load / tension;
                }
                catch (FormatException)
                {
                    MessageBox.Show("Неверный формат!");
                    return false;
                }
            }
            if (groupBox3.Enabled)
            {
                try
                {
                    double koef1 = 0;
                    switch (comboBoxStressUnit.SelectedIndex)
                    {
                        case 0:
                            koef1 = 1;
                            break;
                        case 1:
                            koef1 = _g;
                            break;
                        case 2:
                            koef1 = 10;
                            break;  
                        default:
                            MessageBox.Show("Выберите единицы измерения!");
                            return false;                           
                    }
                    double tension = koef1 * Convert.ToDouble(textBoxStress.Text);

                    double koef2 = 0;
                    switch (comboBoxRelativeLoadUnit.SelectedIndex)
                    {
                        case 0:
                            koef2 = 1;
                            break;
                        case 1:
                            koef2 = _g;
                            break;
                        case 2:
                            koef2 = 10;
                            break;
                        default:
                            MessageBox.Show("Выберите единицы измерения!");
                            return false;                            
                    }
                    double load = koef2 * Convert.ToDouble(textBoxRelativeLoad.Text);
                    _DivisionGammaSigma = load / tension;
                }
                catch (FormatException)
                {
                    MessageBox.Show("Неверный формат!");
                    return false;
                }
            }
            try
            {
                _hscale = Convert.ToDouble(comboBoxHScale.Text.Replace("1:", ""));
                _vscale = Convert.ToDouble(comboBoxVScale.Text.Replace("1:", ""));
                _hgab = Convert.ToDouble(textBoxGabarit.Text);
            }
            catch (FormatException)
            {
                MessageBox.Show("Неверный формат!");
                return false;
            }
            return true;
        }

        private void DrawCatenaryForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Properties.Settings.Default.Save();
        }
    }
}
