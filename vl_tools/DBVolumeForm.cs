using MathParserTK;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace vl_tools
{
    public partial class DBVolumeForm : Form
    {
        public SQLiteConnection connection;

        public DBVolumeForm()
        {
            InitializeComponent();
        }

        private void DBVolumeForm_Load(object sender, EventArgs e)
        {
            try
            {
                
                if (connection != null)
                {
                    //root
                    TreeNode root = this.treeViewDBFolders.Nodes.Add("Сборники");
                    long folder_id = 0;
                    root.Tag = folder_id;
                    //recursive function
                    SQLiteDirBuild(root, 0);
                   
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void SQLiteDirBuild(TreeNode root, long parentId)
        {
            try
            {
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM folders WHERE id IN (SELECT folder_id FROM folders_parent " +
                    "WHERE parent_id=:parent_id) ORDER BY name", connection))
                {
                    cmd.Parameters.Add("parent_id", DbType.Int32).Value = parentId;
                    SQLiteDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        string folderName = reader["name"].ToString();
                        long folder_id = (long)reader["id"];
                        TreeNode nd = root.Nodes.Add(folderName);
                        nd.Tag = folder_id; 
                        SQLiteDirBuild(nd, folder_id);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void GetPositionsInDepth(DataTable tbl, long parentId, string searchPattern)
        {
            try
            {
                //Get positions
                //string cmndTxt = "SELECT * FROM positions WHERE positions.id IN " +
                //"(SELECT position_id FROM positions_folders where folder_id=:folder_id) " +
                //"AND (positions.name LIKE '%" + searchPattern + "%')";
                //Get positions
                string cmndTxt = "SELECT id, code, name, default_unit, price FROM positions WHERE folder_id=:folder_id " +                
                "AND (name LIKE '%" + searchPattern + "%') ORDER BY id";

                using (SQLiteCommand cmd2 = new SQLiteCommand(cmndTxt, connection))
                {                    
                    cmd2.Parameters.Add("folder_id", DbType.Int32).Value = parentId;
                    SQLiteDataReader reader = cmd2.ExecuteReader();
                    while (reader.Read())
                    {                       
                        DataRow row = tbl.NewRow();
                        row["id"] = reader.GetInt64(0);
                        row["Обоснование"] = reader.GetString(1);
                        row["Наименование"] = reader.GetString(2);
                        row["Ед. изм"] = reader.GetString(3); 
                        row["Стоимость"] = reader.GetDouble(4);
                        tbl.Rows.Add(row);
                        //tbl.Rows.Add(reader["id"], reader["code"], reader["name"], reader["default_unit"], reader["price"]);                        
                    }
                }

                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM folders WHERE id IN (SELECT folder_id FROM folders_parent " +
                    "WHERE parent_id=:parent_id) ORDER BY name", connection))
                {
                    cmd.Parameters.Add("parent_id", DbType.Int32).Value = parentId;
                    SQLiteDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        long folder_id = (long)reader["id"];
                        GetPositionsInDepth(tbl, folder_id, searchPattern);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ReFillPositionsGrid()
        {
            try
            {
                string srchTxt = textBoxSearch.Text;
                string fldrTxt = treeViewDBFolders.SelectedNode.Text;
                long curFldrId = (long)treeViewDBFolders.SelectedNode.Tag;
                DataTable tbl = new DataTable();
                tbl.Columns.Add("id", typeof(long));
                tbl.Columns.Add("Обоснование", typeof(string));
                tbl.Columns.Add("Наименование", typeof(string));
                tbl.Columns.Add("Ед. изм", typeof(string));
                tbl.Columns.Add("Стоимость", typeof(double));
                GetPositionsInDepth(tbl, curFldrId, srchTxt);
                this.dataGridViewDBSelect.DataSource = tbl;
                this.dataGridViewDBSelect.Columns["id"].Visible = false;
                this.dataGridViewDBSelect.MultiSelect = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void treeViewDBFolders_AfterSelect(object sender, TreeViewEventArgs e)
        {
            ReFillPositionsGrid();
        }

        private void textBoxSearch_TextChanged(object sender, EventArgs e)
        {
            ReFillPositionsGrid();
        }

        private void InsertIntoObjectTablePosition()
        {
            try
            {
                int c = this.dataGridViewDBSelect.CurrentRow.Index;
                var n = this.dataGridViewObjectPos.Rows.Add();
                dataGridViewObjectPos.Rows[n].Cells["id"].Value = dataGridViewDBSelect.Rows[c].Cells["id"].Value;
                dataGridViewObjectPos.Rows[n].Cells["code"].Value = dataGridViewDBSelect.Rows[c].Cells["Обоснование"].Value;
                dataGridViewObjectPos.Rows[n].Cells["name"].Value = dataGridViewDBSelect.Rows[c].Cells["Наименование"].Value;
                dataGridViewObjectPos.Rows[n].Cells["unit"].Value = dataGridViewDBSelect.Rows[c].Cells["Ед. изм"].Value;
                dataGridViewObjectPos.Rows[n].Cells["price"].Value = dataGridViewDBSelect.Rows[c].Cells["Стоимость"].Value;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void dataGridViewDBSelect_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            InsertIntoObjectTablePosition();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            InsertIntoObjectTablePosition();
        }

        private void dataGridViewObjectPos_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int ncol = e.ColumnIndex;
                int nrow = e.RowIndex;
                if (dataGridViewObjectPos.Columns[ncol].Name == "Formula")
                {
                    MathParser parser = new MathParser();
                    string s = dataGridViewObjectPos.Rows[nrow].Cells[ncol].Value.ToString();
                    try
                    {
                        double d = parser.Parse(s);
                        dataGridViewObjectPos.Rows[nrow].Cells["count"].Value = d;
                        dataGridViewObjectPos.Rows[nrow].Cells[ncol].Style.ForeColor = Color.Black;
                    }
                    catch (Exception)
                    {
                        dataGridViewObjectPos.Rows[nrow].Cells["count"].Value = 0;
                        dataGridViewObjectPos.Rows[nrow].Cells[ncol].Style.ForeColor = Color.Red;
                    }                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void MoveRowUpToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                //DataGridView dgv = this.dataGridViewObjectPos;
                //int irow = dgv.CurrentCell.RowIndex;
                //int icol = dgv.CurrentCell.ColumnIndex;               
                //if (irow == 0) return;
                //DataRow prevRow = table.Rows[irow];
                //DataRow newRow = table.NewRow();
                //newRow.ItemArray = prevRow.ItemArray;
                //table.Rows.Remove(prevRow);
                //table.Rows.InsertAt(newRow, irow - 1);
                //ReloadTables();
                //dgv.CurrentCell = dgv.Rows[irow - 1].Cells[icol];
               
            }
            catch (Exception)
            {

            }
        }
    }
}
