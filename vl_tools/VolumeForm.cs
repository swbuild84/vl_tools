using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using LEP;

namespace vl_tools
{
    public partial class VolumeForm : Form
    {   
        private System.Data.DataTable _tblDetails;  
        private List<BlockObject> _selBlockObjects;
        private List<PlineObject> _selPlineObject;
        private string _templatePath;
        private System.Data.DataTable suprts;

        public VolumeForm(List<BlockObject> selBlockObjects, List<PlineObject> selPlineObject, string path)
        {
            // TODO: Complete member initialization
            try
            {
                suprts = new System.Data.DataTable();
                suprts.Columns.Add("Наименование", System.Type.GetType("System.String"));
                suprts.Columns.Add("Количество", System.Type.GetType("System.Int32"));

                this._selBlockObjects = selBlockObjects;
                this._selPlineObject = selPlineObject;
                _templatePath = path;
                ReadData();
                InitializeComponent();
                if (!Directory.Exists(_templatePath)) throw new FileNotFoundException("Каталог " + _templatePath + "не найден!");
                foreach (string dir in Directory.GetDirectories(_templatePath))
                {
                    string folder = new DirectoryInfo(System.IO.Path.GetDirectoryName(dir+"\\")).Name; ;
                    comboBox1.Items.Add(folder);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ReadData()
        {
            //foreach (DwgObject obj in _dwgObjcts)
            //{
            //    //System.Data.DataTable tbl = DwgObject.CreateDataTable("");
            //    if(obj.GetType() == typeof(BlockObject))
            //    {
            //        BlockObject blk = (BlockObject)obj;
            //        //XElement el = blk.ToXElement();
            //    }
            //    else
            //    {
            //        XElement el = PlineObject.Open(obj.ObjId).ToXSpecification();
            //    }       
            //    //XElement el = DwgObject.GetXMLfromCADEntity(obj.ObjId);
            //    //string type = obj.ObjectType;
            //}

            List<string> detailsTable = new List<string>(); //список попавшихся деталей                      
            
            _tblDetails = new System.Data.DataTable();    //таблица данных
            _tblDetails.Columns.Add("НАИМЕНОВАНИЕ", typeof(string));
            _tblDetails.Columns.Add("КОЛИЧЕСТВО", typeof(double));

            NumberFormatInfo provider = new NumberFormatInfo();
            provider.NumberDecimalSeparator = ".";

            //Подсчет опор  
            IEnumerable<string> supports = from obj in _selBlockObjects.AsEnumerable()
                           where obj.ObjectType.Contains("support")
                           select obj.Name;
            var types = supports.GroupBy(t => t);           

            foreach(var type in types)
            {
                string name = type.Key;
                int count = supports.Count(s => s == name);
                suprts.Rows.Add(name, count);
            }


            for (int i = 0; i < _selBlockObjects.Count; i++)
            {                
                //Подсчет деталей
                IEnumerable<DataRow> query = ((from row in _selBlockObjects[i].Table_1.AsEnumerable()
                                select row).Union(from row2 in _selBlockObjects[i].Table_2.AsEnumerable() 
                                                  select row2));
                foreach (DataRow row in query)
                {
                    if (row["sendInSpecification"].ToString() != "True") continue;
                    string name = row["item_name"].ToString();
                    string count = row["item_count"].ToString();
                    
                    double dcount = Convert.ToDouble(count, provider);
                    if (!detailsTable.Contains(name))
                    {
                        detailsTable.Add(name);
                        DataRow newrow = _tblDetails.NewRow();
                        newrow[0] = name;
                        newrow[1] = dcount;
                        _tblDetails.Rows.Add(newrow);
                    }
                    else
                    {
                        int rowFnd = detailsTable.IndexOf(name);                        
                        double curCount = Convert.ToDouble(_tblDetails.Rows[rowFnd][1], provider);
                        _tblDetails.Rows[rowFnd][1] = curCount + dcount;
                    }
                }                
            }

            for (int i = 0; i < _selPlineObject.Count; i++)
            {
                //Подсчет провода
                string name = _selPlineObject[i].Name;
                double len = _selPlineObject[i].Length;
                if (!detailsTable.Contains(name))
                {
                    detailsTable.Add(name);
                    DataRow newrow = _tblDetails.NewRow();
                    newrow[0] = name;
                    newrow[1] = len;
                    _tblDetails.Rows.Add(newrow);
                }
                else
                {
                    int rowFnd = detailsTable.IndexOf(name);
                    double curCount = Convert.ToDouble(_tblDetails.Rows[rowFnd][1], provider);
                    _tblDetails.Rows[rowFnd][1] = curCount + len;
                }
            }
           
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string folder = comboBox1.Text;
            if (folder == "") return;
            this.listBox1.Items.Clear();
            string sDir = _templatePath + "\\" + folder;
            foreach (string file in Directory.GetFiles(sDir, "*.xlsm"))
            {
                string filename = file.Replace(sDir + "\\", "");
                listBox1.Items.Add(filename);
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private bool IsWorkSheetExist(Microsoft.Office.Interop.Excel.Workbook ObjWorkBook,  string sheetName)
        {
            foreach(object sheetObj in ObjWorkBook.Sheets)
            {
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet) sheetObj;
                if (sheet.Name == sheetName) return true;
            }
            return false;
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            string fileName = listBox1.Text;
            string file = _templatePath + "\\" + comboBox1.Text + "\\" + listBox1.Text;
            //create excel 
            Microsoft.Office.Interop.Excel.Application ObjExcel;
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            
            try
            {
                ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Database dbCurrent = HostApplicationServices.WorkingDatabase;
                string curDwgPath = ed.Document.Name;
                string workFolder = Path.GetDirectoryName(curDwgPath);
                string newXlsFile = workFolder + "\\" + fileName;

                //if file exist?
                if (File.Exists(newXlsFile))
                {
                    DialogResult resInt = MessageBox.Show("Файл " + newXlsFile + " существует. Заменить?", "vl_tools", MessageBoxButtons.YesNo);
                    if (resInt == DialogResult.Yes)
                    {
                        //replace file;
                        File.Copy(file, newXlsFile, true);
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    //copy no perlace
                    File.Copy(file, newXlsFile, false);
                }
                //Книга.
                ObjWorkBook = ObjExcel.Workbooks.Open(newXlsFile);
                //Таблица.
                if (IsWorkSheetExist(ObjWorkBook, "Исходные данные"))
                {
                    ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets["Исходные данные"];
                    int i = 3;
                    foreach (DataRow row in _tblDetails.AsEnumerable())
                    {
                        string name = row.Field<string>(0);
                        double count = row.Field<double>(1);
                        (ObjWorkSheet.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value = name;
                        (ObjWorkSheet.Cells[i, 7] as Microsoft.Office.Interop.Excel.Range).Value = count;
                        i++;
                    }
                }
                //Заполняем лист опоры
                if (IsWorkSheetExist(ObjWorkBook, "Опоры"))
                {
                    ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets["Опоры"];                    
                    int i = 2;
                    foreach (DataRow row in suprts.AsEnumerable())
                    {
                        string name = row.Field<string>(0);
                        int count = row.Field<int>(1);
                        (ObjWorkSheet.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value = name;
                        (ObjWorkSheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value = count;
                        i++;
                    }                    
                }
                //Dwg Properties
                DatabaseSummaryInfo dbInfo = dbCurrent.SummaryInfo;
                DatabaseSummaryInfoBuilder dbInfoBldr = new DatabaseSummaryInfoBuilder(dbInfo);

                if (IsWorkSheetExist(ObjWorkBook, "Штамп"))
                {
                    ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets["Штамп"];
                    ObjWorkSheet.Range["G2"].Value = dbInfoBldr.Title + dbInfoBldr.HyperlinkBase + ".ВОР";
                    ObjWorkSheet.Range["G3"].Value = dbInfoBldr.Comments;
                    ObjWorkSheet.Range["G4"].Value = dbInfoBldr.Subject;
                    ObjWorkSheet.Range["B2"].Value = dbInfoBldr.Author;
                    ObjWorkSheet.Range["B4"].Value = dbInfoBldr.Keywords;
                    IDictionary custProps = dbInfoBldr.CustomPropertyTable;
                    string date = (string)custProps["Дата"];
                    ObjWorkSheet.Range["E2"].Value = date;
                    //ObjWorkSheet.Range["X47"].Value = cntr.ToString() + ".1";
                    //ObjWorkSheet.Range["Z97"].Value = cntr.ToString() + ".2";
                    //ObjWorkSheet.Range["C3"].Value = "Hello";
                }
                ObjWorkBook.Save();
                ObjWorkBook.Close();
                //Собираем мусор
                ObjExcel = null;
                ObjWorkBook = null;
                ObjWorkSheet = null;
                GC.Collect();
                this.DialogResult = DialogResult.OK;
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
            finally
            {

            }
        }
    }
}
