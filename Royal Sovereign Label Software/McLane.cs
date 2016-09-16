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
    public partial class McLane : Form
    {
        protected bool isFormClosing = false;
        private const int WM_CLOSE = 16;
        private const int CARTON = 1;
        private const int PALLET = 2;
        private int labelType = CARTON;
        private OleDbDataReader reader;
        private OleDbConnection con;


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

        private void rb1_CheckedChanged(object sender, EventArgs e)
        {
            labelType = CARTON;
        }

        private void rb2_CheckedChanged(object sender, EventArgs e)
        {
            labelType = PALLET;
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
                MessageBox.Show("Please enter numbers only");
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
                    MessageBox.Show("Please enter numbers only");
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

        private void getData(string sql)
        {
            con = null;
            try
            {
                string connectionString = "Provider=SQLOLEDB.1;Data Source=10.0.0.12\\sqlexpress;Integrated Security=SSPI";
                con = new OleDbConnection(connectionString);
                con.Open();

                OleDbCommand cmd = new OleDbCommand(sql, con);
                reader = cmd.ExecuteReader();
                // Data is accessible through the DataReader object here.
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("No Matching Sales Order Number");
            }
        }

        private void b1_Click(object sender, EventArgs e)
        {

            if (tb_validationcheck() == false)
                return;

            string startingSO, endingSO, output = "";
            startingSO = textBox1.Text;
            endingSO = textBox2.Text;
            startingSO = startingSO.PadLeft(7, '0');
            endingSO = endingSO.PadLeft(7, '0');
            int cartonCount;
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

            //if user select Yes in Message Box
            try
            {
                if (textBox2.Text != "")
                {

                    string sql = "SELECT soh.SalesOrderNo, soh.ShipToName, soh.ShipToAddress1, " +
                        "CASE WHEN soh.ShipToAddress2 is null THEN '' ELSE soh.ShipToAddress2 END as 'ShipToAddress2', " +
                        "soh.ShipToCity, soh.ShipToState, soh.ShipToZipCode, " +
                        "soh.ShipVia, soh.CustomerPONo, sod.UDF_SKU, sod.ItemCodeDesc, sod.ItemCode, " +
                        "sod.QuantityOrdered, CASE WHEN cii.UDF_MASTER_CTN_QTY=0 THEN 1 ELSE cii.UDF_MASTER_CTN_QTY END as 'UDF_MASTER_CTN_QTY', " +
                        "CASE WHEN cii.UDF_MASTER_CTN_QTY = 0 THEN sod.QuantityOrdered ELSE CASE WHEN sod.QuantityOrdered<cii.UDF_MASTER_CTN_QTY THEN sod.QuantityOrdered ELSE sod.QuantityOrdered/cii.UDF_MASTER_CTN_QTY END END as 'TotalCartonQty', " +
                        "soh.WarehouseCode " +
                        "FROM (RSI3...SO_SalesOrderHeader soh INNER JOIN RSI3...SO_SalesOrderDetail sod ON soh.SalesOrderNo = sod.SalesOrderNo)" +
                        "INNER JOIN RSI3...CI_Item cii ON sod.ItemCode = cii.ItemCode " +
                        "WHERE sod.ItemCode Not Like '/%' AND soh.CustomerNo='MCLANE' AND soh.SalesOrderNo>='" + startingSO + "' AND soh.SalesOrderNo<='" + endingSO + "'";

                    getData(sql);
                    // Data is accessible through the DataReader object here.

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        output += reader.GetName(i) + "|";
                    }
                    output += "carton#\n";

                    while (reader.Read())
                    {
                        cartonCount = 1;

                        if (radioButton1.Checked)
                        {
                            while (cartonCount <= reader.GetDecimal(13))
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    output += reader.GetValue(i).ToString() + "|";
                                }
                                output += cartonCount + "\n";
                                cartonCount++;
                            }
                        }
                        else if (radioButton2.Checked)
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                output += reader.GetValue(i).ToString() + "|";
                            }
                            output += cartonCount + "\n";
                        }
                    }
                    con.Close();
                    Console.WriteLine(output);
                    System.IO.File.WriteAllText(@"//10.0.0.10/Public/IT/Label_Printing/bartender_edi/McLaneCartonLabel.csv", output);
                }
                //if textbox2 is null
                else
                {
                    string sql = "SELECT soh.SalesOrderNo, soh.ShipToName, soh.ShipToAddress1, " +
                        "CASE WHEN soh.ShipToAddress2 is null THEN '' ELSE soh.ShipToAddress2 END as 'ShipToAddress2', " +
                        "soh.ShipToCity, soh.ShipToState, soh.ShipToZipCode, " +
                        "soh.ShipVia, soh.CustomerPONo, sod.UDF_SKU, sod.ItemCodeDesc, sod.ItemCode, " +
                        "sod.QuantityOrdered, CASE WHEN cii.UDF_MASTER_CTN_QTY=0 THEN 1 ELSE cii.UDF_MASTER_CTN_QTY END as 'UDF_MASTER_CTN_QTY', " +
                        "CASE WHEN cii.UDF_MASTER_CTN_QTY = 0 THEN sod.QuantityOrdered ELSE CASE WHEN sod.QuantityOrdered<cii.UDF_MASTER_CTN_QTY THEN sod.QuantityOrdered ELSE sod.QuantityOrdered/cii.UDF_MASTER_CTN_QTY END END as 'TotalCartonQty', " +
                        "soh.WarehouseCode " +
                        "FROM (RSI3...SO_SalesOrderHeader soh INNER JOIN RSI3...SO_SalesOrderDetail sod ON soh.SalesOrderNo = sod.SalesOrderNo)" +
                        "INNER JOIN RSI3...CI_Item cii ON sod.ItemCode = cii.ItemCode " +
                        "WHERE sod.ItemCode Not Like '/%' AND soh.CustomerNo='MCLANE' AND soh.SalesOrderNo='" + startingSO + "'";


                    getData(sql);
                    // Data is accessible through the DataReader object here.

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        output += reader.GetName(i) + "|";
                    }
                    output += "carton#|\n";

                    while (reader.Read())
                    {
                        cartonCount = 1;

                        if (radioButton1.Checked)
                        {
                            while (cartonCount <= reader.GetDecimal(13))
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    output += reader.GetValue(i).ToString() + "|";
                                }
                                output += cartonCount + "\n";
                                cartonCount++;
                            }
                        }
                        else if (radioButton2.Checked)
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                output += reader.GetValue(i).ToString() + "|";
                            }
                            output += cartonCount + "\n";
                        }
                    }
                    con.Close();
                    Console.WriteLine(output);
                    System.IO.File.WriteAllText(@"//10.0.0.10/Public/IT/Label_Printing/bartender_edi/McLaneCartonLabel.csv", output);
                }
                textBox1.Clear();
                textBox2.Clear();
                textBox1.Focus();
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("No Matching Sales Order Number");
            }

        }
    }
}
