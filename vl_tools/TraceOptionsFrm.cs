using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace vl_tools
{
    public partial class TraceOptionsFrm : Form
    {
        string _dbPath = "";
        string _GroupName = "";
        public string GroupName
        {
            get 
            { 
                return this.comboBox1.Text; 
            }            
        }
        
        public TraceOptionsFrm(string dbPath, string TraceGroup)
        {
            _dbPath = dbPath;
            _GroupName = TraceGroup;
            InitializeComponent();
        }

        private void TraceOptionsFrm_Load(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            try
            {
                //Init database frame                
                System.Data.DataTable table = new System.Data.DataTable();
                table.Columns.Add(new System.Data.DataColumn("GROUP_NAME", typeof(string), "", MappingType.Attribute));
                table.Columns.Add(new System.Data.DataColumn("ANGLE", typeof(string), "", MappingType.Attribute));
                table.Columns.Add(new System.Data.DataColumn("XMLNAME", typeof(string), "", MappingType.Attribute));
                table.Columns.Add(new System.Data.DataColumn("BLOCKNAME", typeof(string), "", MappingType.Attribute));
                ds.Tables.Add(table);               
                
                //ds.WriteXml(DBPath);                
                if (!File.Exists(_dbPath)) throw new FileNotFoundException("Файл " + _dbPath + " не найден!");
                ds.ReadXml(_dbPath, XmlReadMode.IgnoreSchema);
                var rows = ds.Tables[0].AsEnumerable();
                var groups = rows.GroupBy(r => r.Field<string>("GROUP_NAME"));                
                foreach (var group in groups)
                {
                    string grName = group.Key;
                    this.comboBox1.Items.Add(grName);
                }
                if (_GroupName != "" && this.comboBox1.Items.Contains(_GroupName))
                {
                    this.comboBox1.SelectedIndex = comboBox1.Items.IndexOf(_GroupName);
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
