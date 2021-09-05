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
    public partial class PicketViewerForm : Form
    {
        public PicketViewerForm()
        {
            InitializeComponent();
        }
        public void SetText(string label1, string label2, string label3)
        {
            this.label1.Text = label1;
            this.label2.Text = label2;
            this.label3.Text = label3;
        }

        private void PicketViewerForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Visible=false;
            }           
        }
    }
}
