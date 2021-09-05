using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace SQLTest1
{
    public partial class SQLAbonentsForm : Form
    {
        SqlConnection myConnection = null;
        private BindingSource bindingSource1 = new BindingSource();
        public DataTable table = null;
        public SQLAbonentsForm()
        {
            InitializeComponent();
            myConnection = new SqlConnection("user id=project10;" + "Pwd=project10;" + "Server=ASPSERVER04;" +
    "database=project10;" + "connection timeout=10");
            try
            {
                myConnection.Open();
                SqlCommand myCommand = new SqlCommand("SELECT * FROM district", myConnection);
                SqlDataReader myReader = null;
                myReader = myCommand.ExecuteReader();
                while (myReader.Read())
                {
                    this.comboBox1.Items.Add(myReader["name"].ToString());
                    //MessageBox.Show(myReader["name"].ToString());
                }
                myReader.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string districtName = comboBox1.Text;
            if (districtName == "") return;

            try
            {
                SqlCommand myCommand = new SqlCommand("SELECT * FROM location WHERE district_id= (SELECT district.id FROM district WHERE district.name='" + districtName + "')", myConnection);
                SqlDataReader myReader = null;
                myReader = myCommand.ExecuteReader();
                this.comboBox2.Items.Clear();
                while (myReader.Read())
                {
                    this.comboBox2.Items.Add(myReader["name"].ToString());
                    //MessageBox.Show(myReader["name"].ToString());
                }
                myReader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                myConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string locationName = comboBox2.Text;
            if (locationName == "") return;

            try
            {
                SqlCommand myCommand = new SqlCommand("SELECT * FROM village WHERE location_id= (SELECT location.id FROM location WHERE location.name='" + locationName + "')", myConnection);
                SqlDataReader myReader = null;
                myReader = myCommand.ExecuteReader();
                this.comboBox3.Items.Clear();
                while (myReader.Read())
                {
                    this.comboBox3.Items.Add(myReader["name"].ToString());
                    //MessageBox.Show(myReader["name"].ToString());
                }
                myReader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string village = this.comboBox3.Text;
                if (village == "") return;
                SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT number, name, power, is_ofz from request WHERE village_id = (SELECT village.id FROM village WHERE village.name='" + village + "')", myConnection);
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);
                table = new DataTable();
                table.Locale = System.Globalization.CultureInfo.InvariantCulture;
                dataAdapter.Fill(table);
                bindingSource1.DataSource = table;
                // Resize the DataGridView columns to fit the newly loaded content.
                dataGridView1.AutoResizeColumns(
                    DataGridViewAutoSizeColumnsMode.AllCells);
            }
            catch (Exception)
            {                
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = bindingSource1;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
