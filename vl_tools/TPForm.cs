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

namespace vl_tools
{
    public partial class TPForm : Form
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


        public TPForm()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void TPForm_Load(object sender, EventArgs e)
        {
            try
            {
                //Init database frame
                TPtable.Columns.Add(new DataColumn("POWER", typeof(int), "", MappingType.Attribute));
                TPtable.Columns.Add(new DataColumn("TP_NAME", typeof(string), "", MappingType.Attribute));
                TPtable.Columns.Add(new DataColumn("DWGNAME", typeof(string), "", MappingType.Attribute));
                TPtable.Columns.Add(new DataColumn("BLOCKNAME", typeof(string), "", MappingType.Attribute));
                ds.Tables.Add(TPtable);
                ReadTPData();

                automates = new List<string>();
                string tpName = m_group.Key;
                this.dataGridView2.Rows.Add(8);
                dataGridView2.Rows[0].Cells[0].Value = "Номер ТП";
                dataGridView2.Rows[0].Cells[1].Value = tpName;
                dataGridView2.Rows[1].Cells[0].Value = "Pуст";
                dataGridView2.Rows[2].Cells[0].Value = "k одноврем.";
                dataGridView2.Rows[2].Cells[1].Value = 0.65;
                dataGridView2.Rows[3].Cells[0].Value = "Pрасч";
                dataGridView2.Rows[4].Cells[0].Value = "cos phi";
                dataGridView2.Rows[4].Cells[1].Value = 0.88;
                dataGridView2.Rows[5].Cells[0].Value = "Sрасч";
                dataGridView2.Rows[6].Cells[0].Value = "Мощность тр-ра необходимая";
                dataGridView2.Rows[7].Cells[0].Value = "Загрузка трансформатора, %";
                
                //dataGridView2.Rows[1].Cells[1].Value = tpName;
                //this.labelTP_NUM.Text = tpName;

                dataGridViewLines.ColumnCount = 1;
                dataGridViewLines.ColumnHeadersVisible = true;

                DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
                columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                columnHeaderStyle.Font = new Font("Microsoft Sans Serif", 10);
                columnHeaderStyle.BackColor = Color.SlateGray;

                dataGridViewLines.ColumnHeadersDefaultCellStyle = columnHeaderStyle;
                dataGridView2.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

                // Set the column header names.
                DataGridViewColumn col1 = dataGridViewLines.Columns[0];
                col1.Name = "Наименование";
                col1.Width = 200;
                col1.ReadOnly = true;
                col1.SortMode = DataGridViewColumnSortMode.NotSortable;


                columnCellStyle = new DataGridViewCellStyle();
                columnCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                columnCellStyle.Font = new Font("Microsoft Sans Serif", 8);
                columnCellStyle.WrapMode = DataGridViewTriState.True;               
                  
                this.dataGridViewLines.Rows.Add(14);
                dataGridViewLines.Rows[0].Cells[0].Value = "Сущ. нагрузка, кВт";
                dataGridViewLines.Rows[1].Cells[0].Value = "Кол-во сущ. потребителей";
                dataGridViewLines.Rows[2].Cells[0].Value = "Потребители присоединяемые";
                dataGridViewLines.Rows[3].Cells[0].Value = "Количество присоединяемых потребителей";
                dataGridViewLines.Rows[4].Cells[0].Value = "Присоединяемая мощность, кВт";
                dataGridViewLines.Rows[5].Cells[0].Value = "Коэффициент одновременности";
                dataGridViewLines.Rows[6].Cells[0].Value = "Расчетная мощность";
                dataGridViewLines.Rows[7].Cells[0].Value = "cos Phi";
                dataGridViewLines.Rows[8].Cells[0].Value = "Расчетный ток";
                dataGridViewLines.Rows[9].Cells[0].Value = "Номинал автоматического выключателя";
                
                dataGridViewLines.Rows[10].Cells[0].Value = "Марка и сечение провода";
                dataGridViewLines.Rows[11].Cells[0].Value = "Длина магистрали, м";
                dataGridViewLines.Rows[12].Cells[0].Value = "Потери напряжения, %";
                dataGridViewLines.Rows[13].Cells[0].Value = "Минимальный ток однофазного КЗ, А";

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

                double SumPower = 0;
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
                    int newcol = dataGridViewLines.Columns.Add(fiderName, fiderName);
                    dataGridViewLines.Columns[newcol].Width = 100;
                    dataGridViewLines.Columns[newcol].DefaultCellStyle = columnCellStyle;
                    dataGridViewLines.Columns[newcol].SortMode = DataGridViewColumnSortMode.NotSortable;
                    dataGridViewLines.Rows[2].Cells[newcol].Value = longAbons;
                    dataGridViewLines.Rows[3].Cells[newcol].Value = f.Count().ToString();
                    dataGridViewLines.Rows[4].Cells[newcol].Value = power.ToString("0.00");
                    dataGridViewLines.Rows[7].Cells[newcol].Value = 0.88;                    

                    fidercol++;
                    SumPower += power;
                }
                //tblSmall.Cells[1, 1].TextString = SumPower.ToString("0.00");
                dataGridViewLines.AutoResizeColumnHeadersHeight();
                dataGridViewLines.AutoResizeRows();
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

        private void Calc()
        {

            try
            {
                double Pust = 0;
                for (int i = 1; i < dataGridViewLines.Columns.Count; i++)
                {                    
                    int totalcount = Convert.ToInt32(dataGridViewLines.Rows[3].Cells[i].Value) + Convert.ToInt32(dataGridViewLines.Rows[1].Cells[i].Value);
                    double koef = KoefOdnovrem(totalcount);
                    dataGridViewLines.Rows[5].Cells[i].Value = koef;                    
                    double NominalPower = (Convert.ToDouble(dataGridViewLines.Rows[0].Cells[i].Value) + Convert.ToDouble(dataGridViewLines.Rows[4].Cells[i].Value));
                    double RaschetPower= NominalPower * koef;
                    dataGridViewLines.Rows[6].Cells[i].Value = RaschetPower;
                    double kosPhi = Convert.ToDouble(dataGridViewLines.Rows[7].Cells[i].Value);
                    double RaschetTok = RaschetPower / 1.73 / .38/kosPhi;
                    dataGridViewLines.Rows[8].Cells[i].Value= string.Format("{0:0.00}", RaschetTok);
                    dataGridViewLines.Rows[9].Cells[i].Value = GetAutomatNominal(RaschetTok);
                    //get electrical params







                    Pust += NominalPower;
                }
                dataGridView2.Rows[1].Cells[1].Value = string.Format("{0:0.00}", Pust);
                double koefTP= Convert.ToDouble(dataGridView2.Rows[2].Cells[1].Value);
                dataGridView2.Rows[3].Cells[1].Value = string.Format("{0:0.00}", Pust * koefTP); 
                double kos = Convert.ToDouble(dataGridView2.Rows[4].Cells[1].Value);
                double S = Pust * koefTP / kos;
                dataGridView2.Rows[5].Cells[1].Value = string.Format("{0:0.00}", S);
                double TransPower = GetTrans(S);
                dataGridView2.Rows[6].Cells[1].Value = GetTrans(TransPower);
                dataGridView2.Rows[7].Cells[1].Value = string.Format("{0:0.00%}", S/ TransPower);

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
            double[] stdNominals = { 25, 40, 63, 100, 160, 250, 400, 630};
            if(S<=0) return 0;
            if(S>630) return 0;
            for (int i = 0; i < stdNominals.Length; i++)
            {
                if (stdNominals[i] >= S)
                {
                    return stdNominals[i];
                }
            }
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

        private void buttonAddCol_Click(object sender, EventArgs e)
        {
            string fiderName = "";
            //StandardWindows.StandardWindowsElements
            if (StandardWindowsElements.InputBox("Имя нового фидера?", "Имя нового фидера?", ref fiderName) != DialogResult.OK) return;            
            
            int newcol = dataGridViewLines.Columns.Add(fiderName, fiderName);
            dataGridViewLines.Columns[newcol].Width = 100;
            dataGridViewLines.Columns[newcol].DefaultCellStyle = columnCellStyle;
            dataGridViewLines.Columns[newcol].SortMode = DataGridViewColumnSortMode.NotSortable;
            //dataGridView1.Rows[2].Cells[newcol].Value = longAbons;
            //dataGridView1.Rows[3].Cells[newcol].Value = f.Count().ToString();
            //dataGridView1.Rows[4].Cells[newcol].Value = power.ToString("0.00");
            dataGridViewLines.Rows[7].Cells[newcol].Value = 0.88;
        }

        private void buttonRemoveCol_Click(object sender, EventArgs e)
        {
            int curCol = dataGridViewLines.CurrentCell.ColumnIndex;
            if(curCol!=0) dataGridViewLines.Columns.RemoveAt(curCol);
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
                            cmd.Parameters.Add("power", DbType.Double).Value = Convert.ToDouble(comboBoxPower.Text);
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

        private void dataGridViewLines_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //DataGridViewCellStyle newstyle=new DataGridViewCellStyle();
            //newstyle.BackColor = Color.AliceBlue;
            
            //dataGridViewLines.Rows[e.RowIndex].Cells[e.ColumnIndex].Style = newstyle;            
        }
    }
}
