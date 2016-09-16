using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Royal_Sovereign_Label_Software
{
    public partial class McLane : Form
    {
        protected bool isFormClosing = false;
        private const int WM_CLOSE = 16;
        //private OleDbDataReader reader;
        //private OleDbConnection con;


        public McLane()
        {
            InitializeComponent();
            radioButton1.Checked = true;
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_CLOSE)
                isFormClosing = true;
            base.WndProc(ref m);
        }

        private void tb1_EnterPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
                textBox2.Focus();
        }

        private void tb2_EnterPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
                button1.Focus();
        }

        private bool tb_validationcheck()
        {
            int numCheck;
            if (!int.TryParse(textBox1.Text, out numCheck))
            {
                MessageBox.Show("Please enter valid starting SO#");
                textBox1.SelectAll();
                textBox1.Focus();
                return false;
            }
            else if (textBox1.Text.Length > 7 || textBox1.Text.Length == 0)
            {
                MessageBox.Show("Please enter valid starting SO#");
                textBox1.SelectAll();
                textBox1.Focus();
                return false;
            }
            else if (!int.TryParse(textBox2.Text, out numCheck))
            {
                if (textBox2.Text != "")
                {
                    MessageBox.Show("Please enter valid starting SO#");
                    textBox2.SelectAll();
                    textBox2.Focus();
                    return false;
                }
            }
            else if (textBox2.Text.Length > 7)
            {
                MessageBox.Show("Please enter valid ending SO#");
                textBox2.SelectAll();
                textBox2.Focus();
                return false;
            }
            return true;
        }

        //private void getData(string sql)
        //{
        //    con = null;
        //    try
        //    {
        //        string connectionString = "Provider=SQLOLEDB.1;Data Source=10.0.0.12\\sqlexpress;Integrated Security=SSPI";
        //        con = new OleDbConnection(connectionString);
        //        con.Open();

        //        OleDbCommand cmd = new OleDbCommand(sql, con);
        //        reader = cmd.ExecuteReader();
        //        // Data is accessible through the DataReader object here.
        //    }
        //    catch (InvalidOperationException)
        //    {
        //        MessageBox.Show("No Matching Sales Order Number");
        //    }
        //}

        private void b1_Click(object sender, EventArgs e)
        {

            if (tb_validationcheck() == false)
                return;

            string startingSO, endingSO;
            string[] SOs = new string[2];
            startingSO = textBox1.Text;
            endingSO = textBox2.Text;
            startingSO = startingSO.PadLeft(7, '0');
            endingSO = endingSO.PadLeft(7, '0');

            DialogResult dr;

            //MessageBox to confirm printing
            if (textBox2.Text == "")
                dr = MessageBox.Show("Are you sure to print labels for " + startingSO + "?", "McLane", MessageBoxButtons.YesNo);
            else
                dr = MessageBox.Show("Are you sure to print labels between " + startingSO + " and " + endingSO + "?", "McLane", MessageBoxButtons.YesNo);

            //If user selected No in MessageBox, halt
            if (dr == DialogResult.No)
            {
                textBox1.Focus();
                return;
            }

            //if user selected yes in MessageBox
            SOs[0] = startingSO;
            SOs[1] = endingSO;
            ThreadMethods lm = new ThreadMethods();
            Thread printThread;
            if (radioButton1.Checked)
                printThread = new Thread(lm.PrintMcLaneCartonLabel);
            else
                printThread = new Thread(lm.PrintMcLanePalletLabel);

            printThread.Start(SOs);

            textBox1.Clear();
            textBox2.Clear();
            textBox1.Focus();

        }
    }
}
