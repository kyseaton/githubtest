using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace NewBTASProto
{
    public partial class frmVECustomerBats : Form
    {

        DataSet Bats = new DataSet();
        int max;

        public frmVECustomerBats()
        {
            InitializeComponent();
            LoadData();
            bindingNavigator1.BindingSource = bindingSource1;
        }
        private void LoadData()
        {
            #region setup the binding

            // The xml to bind to.
            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
            strAccessSelect = @"SELECT * FROM Batteries ORDER BY BatterySerialNumber ASC";

            Bats.Clear();
            OleDbConnection myAccessConn = null;
            // try to open the DB
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(Bats);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                return;
            }
            finally
            {
                myAccessConn.Close();
            }




            // Set the DataSource to the DataSet, and the DataMember
            bindingSource1.DataSource = null;
            bindingSource1.DataSource = Bats;

            bindingSource1.DataMember = "Table";

            comboBox1.DataBindings.Add("Text", bindingSource1, "CustomerName");
            comboBox2.DataBindings.Add("Text", bindingSource1, "BatteryModel");
            textBox3.DataBindings.Add("Text", bindingSource1, "BatterySerialNumber");
            textBox4.DataBindings.Add("Text", bindingSource1, "BatteryBCN");
            textBox5.DataBindings.Add("Text", bindingSource1, "BatteryGroup");


            #endregion

            #region setup the combo boxes
            ComboBox SerNumCB = toolStripCBSerNum.ComboBox;
            SerNumCB.DisplayMember = "BatterySerialNumber";
            SerNumCB.DataSource = bindingSource1;

            List<string> Customers =  Bats.Tables[0].AsEnumerable().Select(x => x[5].ToString()).Distinct().ToList();
            Customers.Sort();
            Customers.Insert(0, "");
            ComboBox CustCB = toolStripCBCustomers.ComboBox;
            //SerNumCB.DisplayMember = "BatterySerialNumber";
            CustCB.DataSource = Customers;

            List<string> Mods = Bats.Tables[0].AsEnumerable().Select(x => x[1].ToString()).Distinct().ToList();
            Mods.Sort();
            Mods.Insert(0, "");
            ComboBox ModCB = toolStripCBBatMod.ComboBox;
            //SerNumCB.DisplayMember = "BatterySerialNumber";
            ModCB.DataSource = Mods;
            #endregion

            max = Bats.Tables[0].AsEnumerable().Max(r => r.Field<int>("BID"));

            foreach (string x in Customers)
            {
                comboBox1.Items.Add(x);
            }

            foreach (string x in Mods)
            {
                comboBox2.Items.Add(x);
            }




        }

        private void bindingSource1_DataError(object sender, BindingManagerDataErrorEventArgs e)
        {
            //here!
        }

        private void bindingSource1_DataMemberChanged(object sender, EventArgs e)
        {

        }

        private void bindingSource1_AddingNew(object sender, AddingNewEventArgs e)
        {

        }

        private void bindingSource1_BindingComplete(object sender, BindingCompleteEventArgs e)
        {

        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void bindingSource1_CurrentItemChanged(object sender, EventArgs e)
        {

        }

        private void bindingSource1_DataSourceChanged(object sender, EventArgs e)
        {

        }

        private void bindingSource1_ListChanged(object sender, ListChangedEventArgs e)
        {

        }

        private void bindingSource1_PositionChanged(object sender, EventArgs e)
        {

        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_Leave(object sender, EventArgs e)
        {

            // Maybe down the line...

            /*            ulong num;

                        if (textBox5.Text.Length == 11 && ulong.TryParse(textBox5.Text, out num))
                        {

                            string pn = textBox5.Text;

                            textBox5.Text = String.Format("({0}) {1}-{2}", pn.Substring(0, 3), pn.Substring(3, 3), pn.Substring(6));

                        }
                        if (textBox5.Text.Length == 11 && ulong.TryParse(textBox5.Text, out num))
                        {

                            string pn = textBox5.Text;

                            textBox5.Text = String.Format("({0}) {1}-{2}", pn.Substring(0, 3), pn.Substring(3, 3), pn.Substring(6));

                        }

                        else
                        {

                            MessageBox.Show("Invalid phone number, please change");

                            textBox5.Focus();


                        }
             *  * */
        }

        private void toolStripCBCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {

            #region setup the binding

            // The xml to bind to.
            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            if (toolStripCBCustomers.Text == "" && toolStripCBBatMod.Text == "")
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                strAccessSelect = @"SELECT * FROM Batteries ORDER BY BatterySerialNumber ASC";
            }

            else if (toolStripCBBatMod.Text == "")
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                strAccessSelect = @"SELECT * FROM Batteries WHERE CustomerName='" + toolStripCBCustomers.Text + "' ORDER BY BatterySerialNumber ASC";
            }
            else if (toolStripCBCustomers.Text == "")
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                strAccessSelect = @"SELECT * FROM Batteries WHERE BatteryModel='" + toolStripCBBatMod.Text + "' ORDER BY BatterySerialNumber ASC";
            }
            else
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                strAccessSelect = @"SELECT * FROM Batteries WHERE CustomerName='" + toolStripCBCustomers.Text + "' AND " + "BatteryModel='" + toolStripCBBatMod.Text + "' ORDER BY BatterySerialNumber ASC";
            }
            
            Bats.Clear();
            OleDbConnection myAccessConn = null;
            // try to open the DB
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(Bats);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                return;
            }
            finally
            {
                myAccessConn.Close();
            }


            #endregion

            #region setup the combo boxes
            ComboBox SerNumCB = toolStripCBSerNum.ComboBox;
            SerNumCB.DisplayMember = "BatterySerialNumber";
            SerNumCB.DataSource = bindingSource1;

            #endregion
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Are you sure you want to remove this Battery?", "Delete Record", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    // set up the db Connection
                    string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                    OleDbConnection conn = new OleDbConnection(connectionString);
                    conn.Open();

                    //get the current row
                    DataRowView current = (DataRowView)bindingSource1.Current;

                    // first test to see if the record already is actually in the database

                    if (current["BID"].ToString() != "")
                    {
                        //record already exist as we need to do an update

                        string cmdStr = "DELETE FROM Batteries WHERE BID=" + current["BID"].ToString();
                        OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                        cmd.ExecuteNonQuery();

                        // Also update the binding source
                        Bats.Tables[0].Rows[bindingNavigator1.BindingSource.Position].Delete();

                    }
                    else
                    {
                        MessageBox.Show("That record was not in the DB. You must save it in order to delete it.");
                    }
                    conn.Close();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Deletion Error" + ex.ToString());
            }
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {

                if (radioButton1.Checked == true) textBox5.Text = "STD";
                else if (radioButton2.Checked == true) textBox5.Text = "Custom";
                else textBox5.Text = "";

                // set up the db Connection
                string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                OleDbConnection conn = new OleDbConnection(connectionString);
                conn.Open();

                //MAKE SURE YOU SELECT THE CURRENT ROW FOR DOUBLE SAVES!!!!!!!!!!!!!!!!!

                //get the current row
                DataRowView current = (DataRowView)bindingSource1.Current;

                // first test to see if the record already is in the database

                //string cmdStr = "Select count(*) from CUSTOMERS where CustomerID=" + current["CustomerID"].ToString(); //get the existence of the record as count
                //OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                //int count = (int)cmd.ExecuteScalar();

                if (current["BID"].ToString() != "")
                {
                    //record already exist as we need to do an update

                    string cmdStr = "UPDATE Batteries SET CustomerName='" + comboBox1.Text +
                        "', BatteryModel='" + comboBox2.Text +
                        "', BatterySerialNumber='" + textBox3.Text +
                        "', BatteryBCN='" + textBox4.Text +
                        "', BatteryGroup='" + textBox5.Text +
                        "' WHERE BID=" + current["BID"].ToString();
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    cmd.ExecuteNonQuery();

                }
                else
                {
                    // we need to insert a new record...
                    // find the max value in the CustomerID column so we know what to assign to the new record
                    max++;
                    string cmdStr = "INSERT INTO Batteries (BID, CustomerName, BatteryModel, BatterySerialNumber, BatteryBCN, BatteryGroup) " +
                        "VALUES (" + (max).ToString() + ",'" +
                        comboBox1.Text + "','" +
                        comboBox2.Text + "','" +
                        textBox3.Text + "','" +
                        textBox4.Text + "','" +
                        textBox5.Text + "')";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    cmd.ExecuteNonQuery();

                    // update the dataTable with the new customer ID also..
                    current[0] = max;


                }

                conn.Close();

                // finally figure out if the bat is a standard or custom one
                // set up the db Connection
      
                conn = new OleDbConnection(connectionString);
                conn.Open();

                // see if there is a record in the standard database

                //string cmdStr = "Select count(*) from CUSTOMERS where CustomerID=" + current["CustomerID"].ToString(); //get the existence of the record as count
                //OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                //int count = (int)cmd.ExecuteScalar();

            }// end try
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {

        }

        private void toolStripCBBatMod_SelectedIndexChanged(object sender, EventArgs e)
        {
            #region setup the binding

            // The xml to bind to.
            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            if (toolStripCBCustomers.Text == "" && toolStripCBBatMod.Text == "")
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                strAccessSelect = @"SELECT * FROM Batteries ORDER BY BatterySerialNumber ASC";
            }

            else if (toolStripCBBatMod.Text == "")
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                strAccessSelect = @"SELECT * FROM Batteries WHERE CustomerName='" + toolStripCBCustomers.Text + "' ORDER BY BatterySerialNumber ASC";
            }
            else if (toolStripCBCustomers.Text == "")
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                strAccessSelect = @"SELECT * FROM Batteries WHERE BatteryModel='" + toolStripCBBatMod.Text + "' ORDER BY BatterySerialNumber ASC";
            }
            else
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                strAccessSelect = @"SELECT * FROM Batteries WHERE CustomerName='" + toolStripCBCustomers.Text + "' AND " + "BatteryModel='" + toolStripCBBatMod.Text + "' ORDER BY BatterySerialNumber ASC";
            }

            Bats.Clear();
            OleDbConnection myAccessConn = null;
            // try to open the DB
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(Bats);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                return;
            }
            finally
            {
                myAccessConn.Close();
            }


            #endregion

            #region setup the combo boxes
            ComboBox SerNumCB = toolStripCBSerNum.ComboBox;
            SerNumCB.DisplayMember = "BatterySerialNumber";
            SerNumCB.DataSource = bindingSource1;

            #endregion
        }

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {
            if (textBox5.Text == "STD")
            {
                radioButton2.Checked = false;
                radioButton1.Checked = true;
            }
            else if (textBox5.Text == "Custom")
            {
                radioButton1.Checked = false;
                radioButton2.Checked = true;
            }
        }

        private void bindingNavigator1_Layout(object sender, LayoutEventArgs e)
        {
            if (textBox5.Text == "STD")
            {
                radioButton2.Checked = false;
                radioButton1.Checked = true;
            }
            else if (textBox5.Text == "Custom")
            {
                radioButton1.Checked = false;
                radioButton2.Checked = true;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
