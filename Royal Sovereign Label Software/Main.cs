using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Royal_Sovereign_Label_Software
{
    public partial class Main : Form
    {
        private Sams sams;
        private McLane mclane;

        public Main()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (sams == null)
            {
                sams = new Sams();
                sams.Show();
            }
            else if (!sams.Visible)
            {
                sams = new Sams();
                sams.Show();
            }
            else
            {
                sams.WindowState = FormWindowState.Normal;
                sams.Focus();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (mclane == null)
            {
                mclane = new McLane();
                mclane.Show();
            }
            else if (!mclane.Visible)
            {
                mclane = new McLane();
                mclane.Show();
            }
            else
            {
                mclane.WindowState = FormWindowState.Normal;
                mclane.Focus();
            }
        }

    }
}
