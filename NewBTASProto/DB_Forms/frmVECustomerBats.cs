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


            #endregion

            #region setup the combo boxes
            ComboBox SerNumCB = toolStripCBSerNum.ComboBox;
            SerNumCB.DisplayMember = "BatterySerialNumber";
            SerNumCB.DataSource = bindingSource1;
            

            //  Setup the drop down to contain all customers availible in the customer table

            // Open database containing all the customer names data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
            strAccessSelect = @"SELECT CustomerName FROM CUSTOMERS ORDER BY CustomerName ASC";

            DataSet Custs = new DataSet();
            myAccessConn = null;
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
                myDataAdapter.Fill(Custs);

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

            List<string> Customers =  Custs.Tables[0].AsEnumerable().Select(x => x[0].ToString()).Distinct().ToList();
            Customers.Sort();
            Customers.Insert(0, "");
            ComboBox CustCB = toolStripCBCustomers.ComboBox;
            //SerNumCB.DisplayMember = "BatterySerialNumber";
            CustCB.DataSource = Customers;

            //  Finally, setup the drop down to contain all customers availible in the customer table

            // Open database containing all the customer names data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
            strAccessSelect = @"SELECT BatteryModel FROM BatteriesCustom ORDER BY BatteryModel ASC";

            DataSet BatsList = new DataSet();
            myAccessConn = null;
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
                myDataAdapter.Fill(BatsList);

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

            List<string> Mods = BatsList.Tables[0].AsEnumerable().Select(x => x[0].ToString()).Distinct().ToList();
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


        private void toolStripCBCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateBats();
            toolStripCBCustomers.SelectionLength = 0;
 
        }

        private void UpdateBats()
        {
            #region setup the binding

            // The xml to bind to.
            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";

            if (toolStripCBCustomers.Text == "" && toolStripCBBatMod.Text == "")
            {
                strAccessSelect = @"SELECT * FROM Batteries ORDER BY BatterySerialNumber ASC";
            }

            else if (toolStripCBBatMod.Text == "")
            {
                strAccessSelect = @"SELECT * FROM Batteries WHERE CustomerName='" + toolStripCBCustomers.Text.Replace("'", "''") + "' ORDER BY BatterySerialNumber ASC";
            }
            else if (toolStripCBCustomers.Text == "")
            {
                strAccessSelect = @"SELECT * FROM Batteries WHERE BatteryModel='" + toolStripCBBatMod.Text.Replace("'", "''") + "' ORDER BY BatterySerialNumber ASC";
            }
            else
            {
                strAccessSelect = @"SELECT * FROM Batteries WHERE CustomerName='" + toolStripCBCustomers.Text.Replace("'", "''") + "' AND " + "BatteryModel='" + toolStripCBBatMod.Text.Replace("'", "''") + "' ORDER BY BatterySerialNumber ASC";
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

            string currentID = "";

            if (comboBox1.Text == "" || comboBox2.Text == "" || textBox3.Text == "")
            {
                MessageBox.Show("Please Enter A Customer, Model and Serial Number in order to create a customer battery");
                return;
            }
            try
            {

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

                currentID = current["BID"].ToString();
                if (currentID != "")
                {
                    //record already exist as we need to do an update

                    string cmdStr = "UPDATE Batteries SET CustomerName='" + comboBox1.Text.Replace("'", "''") +
                        "', BatteryModel='" + comboBox2.Text.Replace("'", "''") +
                        "', BatterySerialNumber='" + textBox3.Text.Replace("'", "''") +
                        "', BatteryBCN='" + textBox4.Text.Replace("'", "''") +
                        "' WHERE BID=" + current["BID"].ToString();
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    cmd.ExecuteNonQuery();

                }
                else
                {
                    // we need to insert a new record...
                    // find the max value in the CustomerID column so we know what to assign to the new record
                    max++;
                    string cmdStr = "INSERT INTO Batteries (BID, CustomerName, BatteryModel, BatterySerialNumber, BatteryBCN) " +
                        "VALUES (" + (max).ToString() + ",'" +
                        comboBox1.Text.Replace("'", "''") + "','" +
                        comboBox2.Text.Replace("'", "''") + "','" +
                        textBox3.Text.Replace("'", "''") + "','" +
                        textBox4.Text.Replace("'", "''") + "')";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    cmd.ExecuteNonQuery();

                    // update the dataTable with the new customer ID also..
                    current[0] = max;
                    currentID = max.ToString();


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

            

            UpdateBats();

            //set the current record to this record, if possible...
            try
            {
                int index = bindingSource1.Find("BID", currentID.ToString());
                if (index >= 0)
                {
                    bindingSource1.Position = index;
                }
            }
            catch(Exception ex)
            {
                return;
            }
            




        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            comboBox1.Text = toolStripCBCustomers.Text;
            comboBox2.Text = toolStripCBBatMod.Text;
            return;
        }

        private void toolStripCBBatMod_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateBats();
        }

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {

        }

        private void bindingNavigator1_Layout(object sender, LayoutEventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
