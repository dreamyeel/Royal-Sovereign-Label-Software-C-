using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Royal_Sovereign_Label_Software
{
    public class ThreadMethods
    {
        private OleDbDataReader reader;
        private OleDbConnection con;

        public ThreadMethods()
        {
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

        private string buildSamsString(string startNo, string endNo)
        {
            return "SELECT soh.SalesOrderNo, soh.ShipToName, soh.ShipToAddress1, "
                + "CASE WHEN soh.ShipToAddress2 IS NULL THEN '' ELSE soh.ShipToAddress2 END AS 'ShipToAddress2', "
                + "soh.ShipToCity, soh.ShipToState, soh.ShipToZipCode, soh.CustomerPONo, CASE WHEN sod.UDF_SKU IS NULL "
                + "THEN CASE WHEN sod.CommentText IS NULL THEN '' ELSE sod.CommentText END ELSE sod.UDF_SKU END AS 'CommentText', sod.ItemCodeDesc, sod.ItemCode, sod.QuantityOrdered, "
                + "CASE WHEN cii.UDF_MASTER_CTN_QTY = 0 THEN '1' ELSE cii.UDF_MASTER_CTN_QTY END AS 'UDF_MASTER_CTN_QTY', "
                + "CASE WHEN cii.UDF_MASTER_CTN_QTY = 0 THEN sod.QuantityOrdered ELSE CASE WHEN sod.QuantityOrdered < cii.UDF_MASTER_CTN_QTY THEN sod.QuantityOrdered "
                + "ELSE sod.QuantityOrdered / cii.UDF_MASTER_CTN_QTY END END AS 'TotalCartonQty', sod.WarehouseCode "
                + "FROM RSI3...SO_SalesOrderHeader AS soh INNER JOIN "
                + "RSI3...SO_SalesOrderDetail AS sod ON soh.SalesOrderNo = sod.SalesOrderNo INNER JOIN "
                + "RSI3...CI_Item AS cii ON sod.ItemCode = cii.ItemCode "
                + "WHERE sod.ItemCode Not Like '/%' AND (soh.CustomerNo='SAMS' OR soh.CustomerNo='WALMART') AND soh.SalesOrderNo>='" + startNo + "' AND soh.SalesOrderNo<='" + endNo + "'";
        }
        private string buildSamsString(string startNo)
        {
            return "SELECT soh.SalesOrderNo, soh.ShipToName, soh.ShipToAddress1, "
                + "CASE WHEN soh.ShipToAddress2 IS NULL THEN '' ELSE soh.ShipToAddress2 END AS 'ShipToAddress2', "
                + "soh.ShipToCity, soh.ShipToState, soh.ShipToZipCode, soh.CustomerPONo, CASE WHEN sod.UDF_SKU IS NULL "
                + "THEN CASE WHEN sod.CommentText IS NULL THEN '' ELSE sod.CommentText END ELSE sod.UDF_SKU END AS 'CommentText', sod.ItemCodeDesc, sod.ItemCode, sod.QuantityOrdered, "
                + "CASE WHEN cii.UDF_MASTER_CTN_QTY = 0 THEN '1' ELSE cii.UDF_MASTER_CTN_QTY END AS 'UDF_MASTER_CTN_QTY', "
                + "CASE WHEN cii.UDF_MASTER_CTN_QTY = 0 THEN sod.QuantityOrdered ELSE CASE WHEN sod.QuantityOrdered < cii.UDF_MASTER_CTN_QTY THEN sod.QuantityOrdered "
                + "ELSE sod.QuantityOrdered / cii.UDF_MASTER_CTN_QTY END END AS 'TotalCartonQty', sod.WarehouseCode "
                + "FROM RSI3...SO_SalesOrderHeader AS soh INNER JOIN "
                + "RSI3...SO_SalesOrderDetail AS sod ON soh.SalesOrderNo = sod.SalesOrderNo INNER JOIN "
                + "RSI3...CI_Item AS cii ON sod.ItemCode = cii.ItemCode "
                + "WHERE sod.ItemCode Not Like '/%' AND soh.SalesOrderNo='" + startNo + "'";
        }

        public void PrintSamsCartonLabel(object so)
        {
            int cartonCount;
            string output = "";
            string[] array = (string[]) so;
            try
            {
                if (array[1] != "0000000")
                {

                    
                    getData(buildSamsString(array[0], array[1]));
                    // Data is accessible through the DataReader object here.

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        output += reader.GetName(i) + "|";
                    }
                    output += "carton#|Type|\n";

                    while (reader.Read())
                    {
                        cartonCount = 1;
                        while (cartonCount <= reader.GetDecimal(13))
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                output += reader.GetValue(i).ToString() + "|";
                            }
                            output += cartonCount + "|C|\n";
                            cartonCount++;
                        }
                    }
                    Console.WriteLine(output);
                    System.IO.File.WriteAllText(@"//10.0.0.10/Public/IT/Label_Printing/bartender_edi/SamsCartonLabel.csv", output);
                }
                else
                {
                    getData(buildSamsString(array[0]));
                    // Data is accessible through the DataReader object here.

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        output += reader.GetName(i) + "|";
                    }
                    output += "carton#|Type|\n";

                    while (reader.Read())
                    {
                        cartonCount = 1;
                        while (cartonCount <= reader.GetDecimal(13))
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                output += reader.GetValue(i).ToString() + "|";
                            }
                            output += cartonCount + "|C|\n";
                            cartonCount++;
                        }
                    }
                    con.Close();
                    Console.WriteLine(output);
                    System.IO.File.WriteAllText(@"//10.0.0.10/Public/IT/Label_Printing/bartender_edi/SamsCartonLabel.csv", output);
                }
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("No Matching Sales Order Number");
            }
        }

        public void PrintSamsPalletLabel(object so)
        {
            int cartonCount;
            string output = "";
            string[] array = (string[])so;
            try
            {
                if (array[1] != "0000000")
                {


                    getData(buildSamsString(array[0], array[1]));
                    // Data is accessible through the DataReader object here.

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        output += reader.GetName(i) + "|";
                    }
                    output += "carton#|Type|\n";

                    while (reader.Read())
                    {
                        cartonCount = 1;
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            if (i == 13)
                                output += "1.0000000000|";
                            else
                                output += reader.GetValue(i).ToString() + "|";
                        }
                        output += cartonCount + "|P|\n";
                    }
                    Console.WriteLine(output);
                    System.IO.File.WriteAllText(@"//10.0.0.10/Public/IT/Label_Printing/bartender_edi/SamsCartonLabel.csv", output);
                }
                else
                {
                    getData(buildSamsString(array[0]));
                    // Data is accessible through the DataReader object here.

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        output += reader.GetName(i) + "|";
                    }
                    output += "carton#|Type|\n";

                    while (reader.Read())
                    {
                        cartonCount = 1;
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            if (i == 13)
                                output += "1.0000000000|";
                            else
                                output += reader.GetValue(i).ToString() + "|";
                        }
                        output += cartonCount + "|P|\n";
                    }
                    con.Close();
                    Console.WriteLine(output);
                    System.IO.File.WriteAllText(@"//10.0.0.10/Public/IT/Label_Printing/bartender_edi/SamsCartonLabel.csv", output);
                }
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("No Matching Sales Order Number");
            }
        }

        private string buildMcLaneString(string startNo, string endNo)
        {
            return "SELECT soh.SalesOrderNo, soh.ShipToName, soh.ShipToAddress1, " +
                "CASE WHEN soh.ShipToAddress2 is null THEN '' ELSE soh.ShipToAddress2 END as 'ShipToAddress2', " +
                "soh.ShipToCity, soh.ShipToState, soh.ShipToZipCode, " +
                "soh.ShipVia, soh.CustomerPONo, sod.UDF_SKU, sod.ItemCodeDesc, sod.ItemCode, " +
                "sod.QuantityOrdered, CASE WHEN cii.UDF_MASTER_CTN_QTY=0 THEN 1 ELSE cii.UDF_MASTER_CTN_QTY END as 'UDF_MASTER_CTN_QTY', " +
                "CASE WHEN cii.UDF_MASTER_CTN_QTY = 0 THEN sod.QuantityOrdered ELSE CASE WHEN sod.QuantityOrdered<cii.UDF_MASTER_CTN_QTY THEN sod.QuantityOrdered ELSE sod.QuantityOrdered/cii.UDF_MASTER_CTN_QTY END END as 'TotalCartonQty', " +
                "soh.WarehouseCode " +
                "FROM (RSI3...SO_SalesOrderHeader soh INNER JOIN RSI3...SO_SalesOrderDetail sod ON soh.SalesOrderNo = sod.SalesOrderNo)" +
                "INNER JOIN RSI3...CI_Item cii ON sod.ItemCode = cii.ItemCode " +
                "WHERE sod.ItemCode Not Like '/%' AND soh.CustomerNo='MCLANE' AND soh.SalesOrderNo>='" + startNo + "' AND soh.SalesOrderNo<='" + endNo + "'";
        }
        private string buildMcLaneString(string startNo)
        {
            return "SELECT soh.SalesOrderNo, soh.ShipToName, soh.ShipToAddress1, " +
                "CASE WHEN soh.ShipToAddress2 is null THEN '' ELSE soh.ShipToAddress2 END as 'ShipToAddress2', " +
                "soh.ShipToCity, soh.ShipToState, soh.ShipToZipCode, " +
                "soh.ShipVia, soh.CustomerPONo, sod.UDF_SKU, sod.ItemCodeDesc, sod.ItemCode, " +
                "sod.QuantityOrdered, CASE WHEN cii.UDF_MASTER_CTN_QTY=0 THEN 1 ELSE cii.UDF_MASTER_CTN_QTY END as 'UDF_MASTER_CTN_QTY', " +
                "CASE WHEN cii.UDF_MASTER_CTN_QTY = 0 THEN sod.QuantityOrdered ELSE CASE WHEN sod.QuantityOrdered<cii.UDF_MASTER_CTN_QTY THEN sod.QuantityOrdered ELSE sod.QuantityOrdered/cii.UDF_MASTER_CTN_QTY END END as 'TotalCartonQty', " +
                "soh.WarehouseCode " +
                "FROM (RSI3...SO_SalesOrderHeader soh INNER JOIN RSI3...SO_SalesOrderDetail sod ON soh.SalesOrderNo = sod.SalesOrderNo)" +
                "INNER JOIN RSI3...CI_Item cii ON sod.ItemCode = cii.ItemCode " +
                "WHERE sod.ItemCode Not Like '/%' AND soh.CustomerNo='MCLANE' AND soh.SalesOrderNo='" + startNo + "'";
        }

        public void PrintMcLaneCartonLabel(object so)
        {
            int cartonCount;
            string output = "";
            string[] array = (string[])so;
            try
            {
                if (array[1] != "0000000")
                {


                    getData(buildMcLaneString(array[0], array[1]));
                    // Data is accessible through the DataReader object here.

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        output += reader.GetName(i) + "|";
                    }
                    output += "carton#|Type|\n";

                    while (reader.Read())
                    {
                        cartonCount = 1;
                        while (cartonCount <= reader.GetDecimal(13))
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                output += reader.GetValue(i).ToString() + "|";
                            }
                            output += cartonCount + "|C|\n";
                            cartonCount++;
                        }
                    }
                    Console.WriteLine(output);
                    System.IO.File.WriteAllText(@"//10.0.0.10/Public/IT/Label_Printing/bartender_edi/McLaneCartonLabel.csv", output);
                }
                else
                {
                    getData(buildMcLaneString(array[0]));
                    // Data is accessible through the DataReader object here.

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        output += reader.GetName(i) + "|";
                    }
                    output += "carton#|Type|\n";

                    while (reader.Read())
                    {
                        cartonCount = 1;
                        while (cartonCount <= reader.GetDecimal(13))
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                output += reader.GetValue(i).ToString() + "|";
                            }
                            output += cartonCount + "|C|\n";
                            cartonCount++;
                        }
                    }
                    con.Close();
                    Console.WriteLine(output);
                    System.IO.File.WriteAllText(@"//10.0.0.10/Public/IT/Label_Printing/bartender_edi/McLaneCartonLabel.csv", output);
                }
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("No Matching Sales Order Number");
            }
        }

        public void PrintMcLanePalletLabel(object so)
        {
            int cartonCount;
            string output = "";
            string[] array = (string[])so;
            try
            {
                if (array[1] != "0000000")
                {


                    getData(buildMcLaneString(array[0], array[1]));
                    // Data is accessible through the DataReader object here.

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        output += reader.GetName(i) + "|";
                    }
                    output += "carton#|Type|\n";

                    while (reader.Read())
                    {
                        cartonCount = 1;
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            if (i == 13)
                                output += "1.0000000000|";
                            else
                                output += reader.GetValue(i).ToString() + "|";
                        }
                        output += cartonCount + "|P|\n";
                    }
                    Console.WriteLine(output);
                    System.IO.File.WriteAllText(@"//10.0.0.10/Public/IT/Label_Printing/bartender_edi/McLaneCartonLabel.csv", output);
                }
                else
                {
                    getData(buildMcLaneString(array[0]));
                    // Data is accessible through the DataReader object here.

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        output += reader.GetName(i) + "|";
                    }
                    output += "carton#|Type|\n";

                    while (reader.Read())
                    {
                        cartonCount = 1;
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            if (i == 13)
                                output += "1.0000000000|";
                            else
                                output += reader.GetValue(i).ToString() + "|";
                        }
                        output += cartonCount + "|P|\n";
                    }
                    con.Close();
                    Console.WriteLine(output);
                    System.IO.File.WriteAllText(@"//10.0.0.10/Public/IT/Label_Printing/bartender_edi/McLaneCartonLabel.csv", output);
                }
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("No Matching Sales Order Number");
            }
        }
    }

}
