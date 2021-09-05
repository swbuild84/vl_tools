using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using MathParserTK;

namespace vl_tools
{
    public partial class UserControl1 : UserControl
    {
        private bool _isSelEntity = false;
        private VLBlockObj selObj;

        SQLiteConnection _connection;
        public SQLiteConnection Connection { get => _connection; set => _connection = value; }

        public UserControl1()
        {            
            InitializeComponent();           
        }

        private void UserControl1_Resize(object sender, EventArgs e)
        {
            //this.textBoxSearch.Text = ((UserControl1) sender).Size.ToString();
        }

        internal void ImpliedSelectionChanged(object sender, EventArgs e)
        {
            try
            {
                Document mdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
                Editor editor = mdiActiveDocument.Editor;
                PromptSelectionResult result = editor.SelectImplied();
                if (result.Status != PromptStatus.OK)
                {
                    this.NoSelectedEntMode();
                    this.toolStripStatusLabel1.Text = "Не выбран объект";
                }
                else
                {
                    SelectionSet set = result.Value;
                    if (set.Count != 1)
                    {
                        this.toolStripStatusLabel1.Text = "Выбрано объектов: "+ set.Count.ToString();
                        this.NoSelectedEntMode();
                    }
                    else
                    {
                        ObjectId id = set.GetObjectIds()[0];
                        using (Transaction transaction = mdiActiveDocument.Database.TransactionManager.StartTransaction())
                        {
                            Entity entity = transaction.GetObject(id, OpenMode.ForRead) as Entity;
                            //Selected is BlockRef
                            if (entity.GetType() == typeof(BlockReference))
                            {
                                this.ShowBlockObject(id);
                                //VLDwgObject obj2 = new VLDwgObject(id);
                                //if (obj2.HasExtData)
                                //{
                                //    this.ShowBlockObject(id);
                                //}
                                //else
                                //{
                                //    this.toolStripStatusLabel1.Text = "Объект не имеет доп. данных";
                                //}
                            }
                            //else if (entity.GetType() == typeof(Polyline))
                            //{
                            //    VLDwgObject obj2 = new VLDwgObject(id);
                            //    if (obj2.HasExtData)
                            //    {
                            //        this.ShowPlineObject(id);
                            //    }
                            //}
                            else
                            {
                                this.NoSelectedEntMode();
                                this.toolStripStatusLabel1.Text = "Не выбран объект с объемами";
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                
            }
        }

        private void ShowBlockObject(ObjectId id)
        {
            try
            {                
                selObj = VLBlockObj.Open(id);
                _isSelEntity = true;
                this.statusStrip1.Text = "";
                this.toolStrip2.Enabled = true;
                this.dataGridViewObjectPos.Rows.Clear();
                foreach (DataRow row in selObj.VolumesTable.Rows)
                {
                    var n = this.dataGridViewObjectPos.Rows.Add();
                    dataGridViewObjectPos.Rows[n].Cells["id"].Value = row["id"];
                    dataGridViewObjectPos.Rows[n].Cells["code"].Value = row["code"];
                    dataGridViewObjectPos.Rows[n].Cells["name"].Value = row["name"];
                    dataGridViewObjectPos.Rows[n].Cells["unit"].Value = row["unit"];
                    dataGridViewObjectPos.Rows[n].Cells["price"].Value = row["price"];
                    dataGridViewObjectPos.Rows[n].Cells["count"].Value = row["count"];
                    dataGridViewObjectPos.Rows[n].Cells["Formula"].Value = row["formula"];
                }
                this.dataGridViewObjectPos.Visible = true;
            }
            catch (Exception)
            {
            }

            //VLDwgObject obj = new VLDwgObject(id);
            //if (obj.HasExtData)
            //{

            //}
            //else
            //{
            //    System.Data.DataTable tbl = new System.Data.DataTable();               
            //    bindingSource1.DataSource = tbl;                
            //}

        }

        private void NoSelectedEntMode()
        {
            this.dataGridViewObjectPos.Visible = false;
            _isSelEntity = false;
            this.selObj = null;
            this.toolStrip2.Enabled = false;
            //throw new NotImplementedException();
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {            
            try
            {

                if (_connection != null)
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
            NoSelectedEntMode();
            this.toolStripStatusLabel1.Text = "Не выбран объект";
        }

        private void SQLiteDirBuild(TreeNode root, long parentId)
        {
            try
            {
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM folders WHERE id IN (SELECT folder_id FROM folders_parent " +
                    "WHERE parent_id=:parent_id) ORDER BY name", _connection))
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

        private void GetPositionsInDepth(System.Data.DataTable tbl, long parentId, string searchPattern)
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

                using (SQLiteCommand cmd2 = new SQLiteCommand(cmndTxt, _connection))
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
                    "WHERE parent_id=:parent_id) ORDER BY name", _connection))
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
                System.Data.DataTable tbl = new System.Data.DataTable();
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
            if (!_isSelEntity) return;
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

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (!_isSelEntity) return;
                if (selObj == null) return;

                selObj.VolumesTable.Rows.Clear();
                foreach (DataGridViewRow row in this.dataGridViewObjectPos.Rows)
                {
                    DataRow newRow = selObj.VolumesTable.NewRow();
                    newRow["id"] = row.Cells["id"].Value;
                    newRow["code"] = row.Cells["code"].Value;
                    newRow["name"] = row.Cells["name"].Value;
                    newRow["unit"] = row.Cells["unit"].Value;
                    newRow["price"] = row.Cells["price"].Value;
                    newRow["count"] = (row.Cells["count"].Value != null ? row.Cells["count"].Value : 0);
                    newRow["formula"] = (row.Cells["Formula"].Value != null ? row.Cells["Formula"].Value : 0);
                    selObj.VolumesTable.Rows.Add(newRow);
                }
                selObj.TryToSave();
            }
            catch (Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.ToString());
            }
        }

        private void dataGridViewObjectPos_CellEndEdit_1(object sender, DataGridViewCellEventArgs e)
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
    }  
}
