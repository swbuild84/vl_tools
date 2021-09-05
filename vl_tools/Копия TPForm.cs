using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using StandardWindows;
using System.IO;
using System.Data.SQLite;
using System.Globalization;

namespace vl_tools
{
    public partial class NewTPForm : Form
    {
        public List<int> dgv1Indexes = new List<int>();
        private DataSet ds = new DataSet("DataSet");
        private System.Data.DataTable TPtable = new System.Data.DataTable("TP");
        public SQLiteConnection connection;

        //private BindingSource m_bindingSource1 = new BindingSource();
        //DataTable m_tbl = new DataTable();
        DataGridViewCellStyle columnCellStyle;

        string _SelectedFilePath;
        /// <summary>
        /// Выдает путь к выбранному файлу шаблона dwg
        /// </summary>
        public string SelectedFilePath
        {
            get { return _SelectedFilePath; }            
        }
        string _SelectedBlockName;
        /// <summary>
        /// Выдает имя блока
        /// </summary>
        public string SelectedBlockName
        {
            get { return _SelectedBlockName; }            
        }

        private string nominalTP;       

        public IGrouping<string, Abonent> m_group;
        public List<string> automates;        
        public string templatePath;


        public NewTPForm()
        {
            InitializeComponent();            
        }

        private DataTable RetrieveProvods()
        {
            DataTable table = new DataTable();
            try
            {
                if (connection != null)
                {
                    using (SQLiteDataAdapter adapter = new SQLiteDataAdapter("SELECT name FROM cables", connection))
                    {
                        adapter.Fill(table);                        
                    }
                }
            }                
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return table;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void TPForm_Load(object sender, EventArgs e)
        {
            try
            {
                //Init provod list from db
                this.Provod.DataSource = RetrieveProvods();
                this.Provod.DisplayMember = "name";
                this.Provod.ValueMember = "name";
                //Init database frame
                TPtable.Columns.Add(new DataColumn("POWER", typeof(int), "", MappingType.Attribute));
                TPtable.Columns.Add(new DataColumn("TP_NAME", typeof(string), "", MappingType.Attribute));
                TPtable.Columns.Add(new DataColumn("DWGNAME", typeof(string), "", MappingType.Attribute));
                TPtable.Columns.Add(new DataColumn("BLOCKNAME", typeof(string), "", MappingType.Attribute));
                ds.Tables.Add(TPtable);
                ReadTPData();

                automates = new List<string>();
                string tpName = m_group.Key;

                //Список абонентов данного ТП
                List<Abonent> TPlist = new List<Abonent>();
                foreach (var t in m_group)
                {
                    TPlist.Add(t);
                }

                int fidercol = 1;
                //Группируем абонентов по фидерам
                var TPfiederGroups = from ab in TPlist
                                     orderby ab.FIDER
                                     group ab by ab.FIDER;
                
                //Проход по фидерам
                foreach (IGrouping<string, Abonent> f in TPfiederGroups)
                {
                    //Список абонентов данного фидера
                    string longAbons = "";
                    //суммарная мощность абонентов
                    double power = 0;
                    //MessageBox.Show(f.Key);
                    foreach (var a in f)
                    {
                        longAbons += a.FIO + "; ";
                        power += a.POWER;                       
                    }
                    string fiderName = f.Key;
                    int rowIndex = dataGridViewLines.Rows.Add();
                    DataGridViewRow row = dataGridViewLines.Rows[rowIndex];
                    row.Cells["FiderName"].Value = fiderName;
                    row.Cells["newUsers"].Value = longAbons;
                    row.Cells["NewUsersCount"].Value = f.Count().ToString();
                    row.Cells["NewLoad"].Value = power.ToString("0.00");
                    row.Cells["cos"].Value = 0.88;                    
                    fidercol++;                    
                }
                dataGridViewTP.Rows.Add();
                dataGridViewTP.Rows[0].Cells["tpNumber"].Value= tpName;

                Calc();
            }            
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        private void buttonCalc_Click(object sender, EventArgs e)
        {
            Calc();
        }

        private double GetDoubleFromCell(DataGridViewCell cell)
        {
            try
            {
                if (cell.Value == null) return 0;
                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = ".";
                double res = Convert.ToDouble(cell.Value.ToString().Replace(",", "."), provider);                
                return res;
                
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private int GetIntFromCell(DataGridViewCell cell)
        {
            try
            {
                if (cell.Value == null) return 0;
                NumberFormatInfo provider = new NumberFormatInfo();
                provider.NumberDecimalSeparator = ".";
                int res = Convert.ToInt32(cell.Value.ToString().Replace(",", "."), provider);
                return res;

            }
            catch (Exception)
            {
                return 0;
            }
        }

        //private double GetKZMin(string provod, double lenght, string circuit_diagram)
        //{
        //}

        private double GetProvodSection(string name)
        {
            double res = 0;
            try
            {
                if (connection != null)
                {
                    using (SQLiteCommand cmd = new SQLiteCommand("SELECT phaza_section FROM cables WHERE name>=:name ORDER BY name LIMIT 1", connection))
                    {
                        cmd.Parameters.Add("name", DbType.String).Value = name;
                        res = (double)cmd.ExecuteScalar();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return res;
        }

        private void Calc()
        {

            try
            {
                double Pust = 0;
                double Prasch = 0;
                double cosPower = 0;
                for (int i = 0; i < dataGridViewLines.Rows.Count; i++)
                {
                    DataGridViewRow row = dataGridViewLines.Rows[i];
                    //Проверка полноты введенных данных

                    //Общее число абонентов на линии
                    int totalcount = GetIntFromCell(row.Cells["existUsersCount"]) +
                        GetIntFromCell(row.Cells["NewUsersCount"]);
                    if (totalcount == 0) continue;

                    double koef = 1;
                    if (row.Cells["KoefOdovr"].Tag == null)
                    {
                        //Вычисляем коэффициент
                        koef = KoefOdnovrem(totalcount);
                        row.Cells["KoefOdovr"].Value = koef;
                    }
                    else
                    {
                        //коэффициент введен вручную!
                        koef = GetDoubleFromCell(row.Cells["KoefOdovr"]);
                    }  

                    double NominalPower = (GetDoubleFromCell(row.Cells["existLoad"])
                        + GetDoubleFromCell(row.Cells["NewLoad"]));

                    double RaschetPower= NominalPower * koef;
                    row.Cells["CalcPower"].Value = RaschetPower;

                    double kosPhi = GetDoubleFromCell(row.Cells["cos"]);
                    
                    double RaschetTok = RaschetPower / 1.73 / .38/kosPhi;
                    row.Cells["Amperage"].Value = string.Format("{0:0.00}", RaschetTok);

                    row.Cells["Automat"].Value = GetAutomatNominal(RaschetTok);
                    Pust += NominalPower;
                    Prasch += RaschetPower;
                    cosPower += kosPhi * RaschetPower;

                    //calc deltaU
                    double lenght = GetDoubleFromCell(row.Cells["Lenght"]);
                    if (lenght > 0)
                    {
                        //Принимаем среднюю длину для приложения нагрузки
                        lenght = lenght / 2;
                        string cable = row.Cells["Provod"].Value.ToString();
                        if (cable != "")
                        {
                            double section = GetProvodSection(cable);
                            if (section > 0)
                            {
                                double deltaU = RaschetPower * lenght / (46 * section);
                                row.Cells["deltaU"].Value = deltaU;
                            }
                        }
                        
                    }
                    //calc TKZ
                }
                DataGridViewRow rowTP = dataGridViewTP.Rows[0];
                rowTP.Cells["nominalLoad"].Value = Pust;
                if (Pust == 0) return;
                //Средневзвешенный коэффициент одновременности
                double koefTP = Prasch / Pust;
                rowTP.Cells["koeffOndovr"].Value = koefTP;
                rowTP.Cells["CalcPowerTP"].Value = Pust * koefTP;
                //dataGridViewTP.Rows[1].Cells["CalcPowerTP"].Value = Prasch;
                //Средневзвешенный cos 
                double kos = cosPower / Prasch;
                rowTP.Cells["cosTP"].Value = kos;
                double S = Pust * koefTP / kos;
                rowTP.Cells["FullLoadTP"].Value = S; 
                
                double TransPower = GetTrans(S);
                rowTP.Cells["NeededPower"].Value = TransPower;
                rowTP.Cells["PercentLoad"].Value = string.Format("{0:0.00%}", S / TransPower);

                nominalTP = TransPower.ToString();
                comboBoxPower.SelectedItem = nominalTP;                               

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private double GetTrans(double S)
        {
            try
            {
                if (connection != null)
                {
                    using (SQLiteCommand cmd = new SQLiteCommand("SELECT power FROM transformers WHERE power>=:power ORDER BY power LIMIT 1", connection))
                    {
                        cmd.Parameters.Add("power", DbType.Double).Value = S;                        
                        return vl_tools.Commands.GetDoubleFString(cmd.ExecuteScalar().ToString());                        
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            //double[] stdNominals = { 25, 40, 63, 100, 160, 250, 400, 630};
            //if(S<=0) return 0;
            //if(S>630) return 0;
            //for (int i = 0; i < stdNominals.Length; i++)
            //{
            //    if (stdNominals[i] >= S)
            //    {
            //        return stdNominals[i];
            //    }
            //}
            return 0;
        }

        private string GetAutomatNominal(double raschetTok)
        {
            //ВА51-35
            double[] stdNominals1 = {16, 25, 31.5, 40, 63, 80, 100, 125, 160, 200, 250, 320, 400};
            //ВА57            
            double[] stdNominals2 = { 630 };
            if (raschetTok <= 0 || raschetTok > 630) return "-";
            if (raschetTok > 400)
            {                
                for(int i=0;i<stdNominals2.Length;i++)
                {
                    if(stdNominals2[i]>=raschetTok)
                    {
                        return stdNominals2[i].ToString();
                    }
                }
            }
            else
            {
                for (int i = 0; i < stdNominals1.Length; i++)
                {
                    if (stdNominals1[i] >= raschetTok)
                    {
                        return stdNominals1[i].ToString();
                    }
                }
            }
            return "-";

        }

        private double KoefOdnovrem(int number)
        {
            try
            {
                double koef = 0;
                if (number == 1) koef = 1;
                if (number == 2) koef = 0.73;
                if (number >= 3 && number < 5) koef = 0.62;
                if (number >= 5 && number < 10) koef = 0.5;
                if (number >= 10 && number < 20) koef = 0.38;
                if (number >= 20) koef = 0.29;
                return koef;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void buttonAddRow_Click(object sender, EventArgs e)
        {
            int nRow = dataGridViewLines.Rows.Add();
            dataGridViewLines.Rows[nRow].Cells["cos"].Value = 0.88;
            //string fiderName = "";
            ////StandardWindows.StandardWindowsElements
            //if (StandardWindowsElements.InputBox("Имя нового фидера?", "Имя нового фидера?", ref fiderName) != DialogResult.OK) return;            
            
            //int newcol = dataGridViewLines.Columns.Add(fiderName, fiderName);
            //dataGridViewLines.Columns[newcol].Width = 100;
            //dataGridViewLines.Columns[newcol].DefaultCellStyle = columnCellStyle;
            //dataGridViewLines.Columns[newcol].SortMode = DataGridViewColumnSortMode.NotSortable;
            ////dataGridView1.Rows[2].Cells[newcol].Value = longAbons;
            ////dataGridView1.Rows[3].Cells[newcol].Value = f.Count().ToString();
            ////dataGridView1.Rows[4].Cells[newcol].Value = power.ToString("0.00");
            //dataGridViewLines.Rows[7].Cells[newcol].Value = 0.88;
        }

        private void buttonRemoveRow_Click(object sender, EventArgs e)
        {
            int curRow = dataGridViewLines.CurrentCell.RowIndex;
            dataGridViewLines.Rows.RemoveAt(curRow);
            //int curCol = dataGridViewLines.CurrentCell.ColumnIndex;
            //if(curCol!=0) dataGridViewLines.Columns.RemoveAt(curCol);
        }

        private void buttonInsert_Click(object sender, EventArgs e)
        {
            ExitAndInsert();            
        }

        private void comboBoxPower_SelectedValueChanged(object sender, EventArgs e)
        {
            nominalTP = comboBoxPower.Text;
            this.comboBoxDiagram.Items.Clear();
            if (comboBoxPower.Text != "")
            {
                try
                {
                    if(connection!=null)
                    {
                        using (SQLiteCommand cmd = new SQLiteCommand("SELECT circuit_diagram FROM transformers WHERE power=:power", connection))
                        {
                            cmd.Parameters.Add("power", DbType.Double).Value = vl_tools.Commands.GetDoubleFString(comboBoxPower.Text);
                            SQLiteDataReader reader = cmd.ExecuteReader();
                            while (reader.Read())
                            {
                                this.comboBoxDiagram.Items.Add(reader["circuit_diagram"].ToString());
                            }
                        }
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            TPListRefill();
        }

        private void ExitAndInsert()
        {
            try
            {
                if(listBox1.SelectedIndex<0) return;
                automates.Clear();


                if(dataGridViewLines.Columns[0].DisplayIndex!=0) //Сдвинули столбец названий - ошибка
                {
                    throw new Exception("Столбец названий перемещать нельзя!");
                }

                var q = from c in this.dataGridViewLines.Columns.Cast<DataGridViewColumn>() orderby c.DisplayIndex select c;
                foreach (DataGridViewColumn column in q)
                {
                    if (column.Visible == true)
                    {
                        if(column.DisplayIndex>0) dgv1Indexes.Add(column.Index);
                    }
                }
                
                foreach (int j in dgv1Indexes)
                {
                    automates.Add(dataGridViewLines.Rows[9].Cells[j].Value.ToString());
                }
                //for (int i = 1; i < dataGridView1.Columns.Count; i++)
                //{
                //    automates.Add(dataGridView1.Rows[9].Cells[i].Value.ToString());
                //}


                var selfilename= from row in TPtable.AsEnumerable()
                                 where row.Field<string>("TP_NAME") == listBox1.Text
                                 select row;
                if (selfilename.Count() != 1) throw new InvalidDataException("TP_NAME " + listBox1.Text + " встречается более 1 раза. Исправьте базу данных!");

                DataRow selTProw = selfilename.ElementAt(0);
                this._SelectedFilePath = templatePath + "\\" + selTProw["DWGNAME"].ToString();
                this._SelectedBlockName= selTProw["BLOCKNAME"].ToString();
                
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ReadTPData()
        {
            try
            {
                string DBPath = templatePath + "\\tp.xml";
                if (!File.Exists(DBPath)) throw new FileNotFoundException("Файл " + DBPath + " не найден!");                
                ds.ReadXml(DBPath, XmlReadMode.IgnoreSchema);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void TPListRefill()
        {
            //ReadTPData();
            try
            { 
                var tpnames = from row in TPtable.AsEnumerable()
                              where row.Field<int>("POWER") == Convert.ToInt32(nominalTP)
                              select row.Field<string>("TP_NAME");
                listBox1.Items.Clear();
                foreach (var tpname in tpnames)
                {
                    listBox1.Items.Add(tpname);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }         

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            ExitAndInsert();
        }

        private void dataGridView1_ColumnDisplayIndexChanged(object sender, DataGridViewColumnEventArgs e)
        {
            
        }

        private void comboBoxPower_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridViewLines_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            string headerText = dataGridViewLines.Columns[e.ColumnIndex].Name;
            string[] cols = { "existLoad", "existUsersCount", "NewUsersCount", "NewLoad", "cos", "KoefOdovr" };
            if (!cols.Contains(headerText)) return;
            if (headerText == "existUsersCount" || headerText == "NewUsersCount")
            {
                try
                {
                    DataGridView dgv = sender as DataGridView;
                    if (dgv == null) return;
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "";
                    int newInt;
                    if (dgv.Rows[e.RowIndex].IsNewRow) { return; }

                    string s_val = e.FormattedValue.ToString();
                    if (s_val == "") { return; }

                    if (!int.TryParse(e.FormattedValue.ToString(), NumberStyles.Float,
                        CultureInfo.CreateSpecificCulture("en-US"), out newInt) || newInt < 0)
                    {
                        e.Cancel = true;
                        dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "Введите положительное целое число";
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else
            {
                try
                {
                    DataGridView dgv = sender as DataGridView;
                    if (dgv == null) return;
                    dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "";
                    double newDouble;
                    if (dgv.Rows[e.RowIndex].IsNewRow) { return; }

                    string s_val = e.FormattedValue.ToString();
                    if (s_val == "") { return; }

                    if (!double.TryParse(e.FormattedValue.ToString(), NumberStyles.Float,
                        CultureInfo.CreateSpecificCulture("en-US"), out newDouble) || newDouble < 0)
                    {
                        e.Cancel = true;
                        dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "Введите положительное число";
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void dataGridViewLines_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //string headerText = dataGridViewLines.Columns[e.ColumnIndex].Name;
            
            //if (headerText == "KoefOdovr")
            //{
            //    dataGridViewLines.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag = "edited";
            //}
        }

        private void dataGridViewLines_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            string headerText = dataGridViewLines.Columns[e.ColumnIndex].Name;
            var cell = dataGridViewLines.Rows[e.RowIndex].Cells[e.ColumnIndex];
            if (headerText == "KoefOdovr")
            {
                string s_val = cell.FormattedValue.ToString();
                if (s_val == "")
                {
                    cell.Tag = null;
                    cell.Style.BackColor = Color.White;
                    return;
                }
                double val = double.Parse(dataGridViewLines.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
                if (val == 0)
                {
                    cell.Tag = null;
                    cell.Style.BackColor = Color.White;
                }
                else
                {
                    cell.Tag = "edited";
                    cell.Style.BackColor = Color.Red;
                }
            }
        }
    }
}
