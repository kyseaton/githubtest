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

        // current data
        string curTemp1;
        string curTemp2;
        string curTemp3;
        string curTemp4;

        // we use this bool to allow us to allow the databinding indext to be changed...
        bool Inhibit = true;
        bool InhibitCB = true;

        public frmVECustomerBats()
        {
            InitializeComponent();
            LoadData();
            bindingNavigator1.BindingSource = bindingSource1;

            bindingNavigator1.CausesValidation = true;

        }
        private void LoadData()
        {
            #region setup the binding

            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
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
                MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                lock (Main_Form.dataBaseLock)
                {
                    myAccessConn.Open();
                    myDataAdapter.Fill(Bats);
                    myAccessConn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                return;
            }
            finally
            {

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

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
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
                MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                lock (Main_Form.dataBaseLock)
                {
                    myAccessConn.Open();
                    myDataAdapter.Fill(Custs);
                    myAccessConn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                return;
            }
            finally
            {
                
            }

            List<string> Customers =  Custs.Tables[0].AsEnumerable().Select(x => x[0].ToString()).Distinct().ToList();
            Customers.Sort();
            Customers.Insert(0, "");
            ComboBox CustCB = toolStripCBCustomers.ComboBox;
            //SerNumCB.DisplayMember = "BatterySerialNumber";
            CustCB.DataSource = Customers;

            //  Finally, setup the drop down to contain all customers availible in the customer table

            // Open database containing all the customer names data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
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
                MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                lock (Main_Form.dataBaseLock)
                {
                    myAccessConn.Open();
                    myDataAdapter.Fill(BatsList);
                    myAccessConn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                return;
            }
            finally
            {

            }

            List<string> Mods = BatsList.Tables[0].AsEnumerable().Select(x => x[0].ToString()).Distinct().ToList();
            Mods.Sort();
            Mods.Insert(0, "");
            ComboBox ModCB = toolStripCBBatMod.ComboBox;
            //SerNumCB.DisplayMember = "BatterySerialNumber";
            ModCB.DataSource = Mods;
            #endregion

            try
            {
                max = Bats.Tables[0].AsEnumerable().Max(r => r.Field<int>("BID"));
            }
            catch
            {
                max = 1;
            }


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
            updateCurVals();
            InhibitCB = false;
            lastValid = false;
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        int oldPositionCust = 0;

        private void toolStripCBCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (InhibitCB)
            {
                return;
            }

            //Validate before moving
            if (ValidateIt())
            {
                // move back
                InhibitCB = true;
                toolStripCBCustomers.SelectedIndex = oldPositionCust;
            }
            else
            {
                oldPositionCust = toolStripCBCustomers.SelectedIndex;
                // make sure the add new is enabled..
                bindingNavigatorAddNewItem.Enabled = true;
                UpdateBats();
                updateCurVals();
            }
            InhibitCB = true; 
        }

        private void UpdateBats()
        {
            #region setup the binding

            // The xml to bind to.
            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";

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
                MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                lock (Main_Form.dataBaseLock)
                {
                    myAccessConn.Open();
                    myDataAdapter.Fill(Bats);
                    myAccessConn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                return;
            }
            finally
            {

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
                if (MessageBox.Show(new Form() { TopMost = true }, "Are you sure you want to remove this Battery?", "Delete Record", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    // set up the db Connection
                    string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                    OleDbConnection conn = new OleDbConnection(connectionString);

                    //get the current row
                    DataRowView current = (DataRowView)bindingSource1.Current;

                    // first test to see if the record already is actually in the database

                    if (current["BID"].ToString() != "")
                    {
                        //record already exist as we need to do an update

                        string cmdStr = "DELETE FROM Batteries WHERE BID=" + current["BID"].ToString();
                        OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                        lock (Main_Form.dataBaseLock)
                        {
                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }

                        // Also update the binding source
                        Bats.Tables[0].Rows[bindingNavigator1.BindingSource.Position].Delete();

                    }
                    else
                    {
                        MessageBox.Show(new Form() { TopMost = true }, "That record was not in the DB. You must save it in order to delete it.");
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(new Form() { TopMost = true }, "Deletion Error" + ex.ToString());
            }
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            if (bindingNavigator1.BindingSource.Position == -1)
            {
                string temp1 = textBox3.Text;
                string temp2 = textBox4.Text;
                string temp3 = comboBox1.Text;
                string temp4 = comboBox2.Text;

                bindingNavigator1.BindingSource.AddNew();
                bindingNavigator1.BindingSource.Position = 0;

                textBox3.Text = temp1;
                textBox4.Text = temp2;
                comboBox1.Text = temp3;
                comboBox2.Text = temp4;
            }

            string currentID = "";

            if (comboBox1.Text == "" || comboBox2.Text == "" || textBox3.Text == "")
            {
                MessageBox.Show(new Form() { TopMost = true }, "Please Enter A Customer, Model and Serial Number in order to create a customer battery");
                return;
            }
            // set up the db Connection
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            OleDbConnection conn = new OleDbConnection(connectionString);

            try
            {

                

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
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                    // Also update the serial number in the other workOrders table!
                    cmdStr = "UPDATE WorkOrders SET BatterySerialNumber='" + textBox3.Text.Replace("'", "''") + "' WHERE BatterySerialNumber='" + current["BatterySerialNumber"].ToString() + "'";
                    cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                    // Also update the serial number in the other workOrders table!
                    cmdStr = "UPDATE WorkOrders SET BatteryModel='" + comboBox2.Text.Replace("'", "''") + "' WHERE BatterySerialNumber='" + current["BatterySerialNumber"].ToString() + "'";
                    cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    
                    //now force an update on the binding by moving one ahead and then back...
                    toolStripCBSerNum.ComboBox.Text = textBox3.Text.Replace("'", "''");

                    MessageBox.Show(new Form() { TopMost = true }, "Battery serial number " + textBox3.Text.Replace("'", "''") + "'s entry has been updated.");

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

                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    MessageBox.Show(new Form() { TopMost = true }, "Battery serial number " + textBox3.Text + "'s entry has been created.");

                    // update the dataTable with the new customer ID also..
                    current[0] = max;
                    currentID = max.ToString();


                }


            }// end try
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            UpdateBats();

            bindingNavigatorAddNewItem.Enabled = true;

            // also update the current vals..
            updateCurVals();

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

            //Last step!  We need to also update the pci table so our graphs look good...
            // we'll need to load the data for this model...
            DataSet modelInfo = new DataSet();
            string strAccessSelect = @"SELECT BTECH,VOLT,NCELLS,BCVMIN,BCVMAX,CCVMMIN,CCVMAX,CCAPV FROM BatteriesCustom WHERE BatteryModel='" + comboBox2.Text + "'";
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, conn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                lock (Main_Form.dataBaseLock)
                {
                    conn.Open();
                    myDataAdapter.Fill(modelInfo);
                    conn.Close();
                }

            }
            catch (Exception ex)
            {
                // didn't work...
                // no biggie!
                return;
            }

            // Also update the pci dataTable!
            for (int i = 0; i < 16; i++)
            {
                if (((Main_Form)this.Owner).pci.Rows[i][9].ToString() == textBox3.Text)
                {
                    //update the four rows and then break...
                    ((Main_Form)this.Owner).pci.Rows[i][0] = comboBox2.Text;
                    if (modelInfo.Tables[0].Rows[0][0].ToString() != "") { ((Main_Form)this.Owner).pci.Rows[i][1] = modelInfo.Tables[0].Rows[0][0].ToString(); }        // tech
                    if (modelInfo.Tables[0].Rows[0][1].ToString() != "") { ((Main_Form)this.Owner).pci.Rows[i][2] = modelInfo.Tables[0].Rows[0][1].ToString(); }        // (NomV) negative 1 is the default...
                    if (modelInfo.Tables[0].Rows[0][2].ToString() != "") { ((Main_Form)this.Owner).pci.Rows[i][3] = modelInfo.Tables[0].Rows[0][2].ToString(); }        // (NCells) negative 1 is the default...
                    if (modelInfo.Tables[0].Rows[0][3].ToString() != "") { ((Main_Form)this.Owner).pci.Rows[i][4] = modelInfo.Tables[0].Rows[0][3].ToString(); }        // (BCVMIN) negative 1 is the default...
                    if (modelInfo.Tables[0].Rows[0][4].ToString() != "") { ((Main_Form)this.Owner).pci.Rows[i][5] = modelInfo.Tables[0].Rows[0][4].ToString(); }        // (BCVMAX) 24 is the default...
                    if (modelInfo.Tables[0].Rows[0][5].ToString() != "") { ((Main_Form)this.Owner).pci.Rows[i][6] = modelInfo.Tables[0].Rows[0][5].ToString(); }        // (CCVMMIN) negative 1 is the default...
                    if (modelInfo.Tables[0].Rows[0][6].ToString() != "") { ((Main_Form)this.Owner).pci.Rows[i][7] = modelInfo.Tables[0].Rows[0][6].ToString(); }        // (CCVMAX) 1.75 is the default...
                    if (modelInfo.Tables[0].Rows[0][7].ToString() != "") { ((Main_Form)this.Owner).pci.Rows[i][8] = modelInfo.Tables[0].Rows[0][7].ToString(); }        // (CCAPV) negative 1 is the default...
                }
            }

        }

        private void updateCurVals()
        {
            // update the current vars....
            //current data..
            curTemp1 = textBox3.Text;
            curTemp2 = textBox4.Text;
            curTemp3 = comboBox1.Text;
            curTemp4 = comboBox2.Text;
        }
        


        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            if (toolStripCBSerNum.Text == "")
            {
                comboBox1.Text = toolStripCBCustomers.Text;
                comboBox2.Text = toolStripCBBatMod.Text;
                bindingNavigatorAddNewItem.Enabled = false;
            }
            lastValid = false;

        }

        int oldPositionBat = 0;

        private void toolStripCBBatMod_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (InhibitCB)
            {
                return;
            }

            //Validate before moving
            if (ValidateIt())
            {
                // move back
                InhibitCB = true;
                toolStripCBBatMod.SelectedIndex = oldPositionBat;
            }
            else
            {
                oldPositionBat = toolStripCBBatMod.SelectedIndex;
                bindingNavigatorAddNewItem.Enabled = true;
                UpdateBats();
                updateCurVals();
            }
            InhibitCB = true;
            
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

        bool lastValid = false;

        private void bindingNavigator1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid)
                {
                    Bats.Tables[0].Rows[Bats.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void bindingNavigatorMovePreviousItem_Click(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid)
                {
                    Bats.Tables[0].Rows[Bats.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }

        }

        private void bindingNavigatorMoveFirstItem_Click(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid)
                {
                    Bats.Tables[0].Rows[Bats.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void toolStripCBSerNum_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid)
                {
                    Bats.Tables[0].Rows[Bats.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }

        }


        private bool ValidateIt()
        {
            // do we need to validate?
            if (textBox3.Text != curTemp1 ||
                textBox4.Text != curTemp2 ||
                comboBox1.Text != curTemp3 ||
                comboBox2.Text != curTemp4 )
            {
                // they don't match!
                // ask if the user is sure that they want to continue...
                DialogResult dialogResult = MessageBox.Show(this, "Looks like this record has been updated without being saved.  Are you sure you want to navigate away without saving?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                {
                    lastValid = false;
                    return true;
                }
            }
            lastValid = true;
            return false;
        }


        private void bindingNavigator1_ItemClicked_1(object sender, ToolStripItemClickedEventArgs e)
        {
            Inhibit = false;
            bindingNavigator1.Focus();
        }

        private void frmVECustomerBats_FormClosing_1(object sender, FormClosingEventArgs e)
        {
            //Validate before moving
            if (ValidateIt())
            {
                Inhibit = true;
                // move back
                e.Cancel = true;
            }
            else
            {
                Inhibit = true;
            }
        }

        private void frmVECustomerBats_Shown(object sender, EventArgs e)
        {
            updateCurVals();
        }

        private void bindingNavigatorMoveNextItem_Click(object sender, EventArgs e)
        {

        }

        private void bindingNavigator1_Validating_1(object sender, CancelEventArgs e)
        {
            if (Inhibit) { return; }

            //Validate before moving
            if (ValidateIt())
            {
                Inhibit = true;
                // move back
                e.Cancel = true;

            }
            else
            {
                Inhibit = true;
                InhibitCB = true;
            }

        }

        private void toolStripCBSerNum_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (InhibitCB) 
            {
                return; 
            }

            //Validate before moving
            if (ValidateIt())
            {
                
                // move back
                InhibitCB = true;
                toolStripCBSerNum.SelectedIndex = bindingNavigator1.BindingSource.Position;

            }
            else
            {
                updateCurVals();
            }
            InhibitCB = true;
        }

        private void toolStripCBSerNum_Validating(object sender, CancelEventArgs e)
        {

        }

        private void toolStripCBSerNum_DropDown(object sender, EventArgs e)
        {
            InhibitCB = false;
        }

        private void toolStripCBCustomers_DropDown(object sender, EventArgs e)
        {
            InhibitCB = false;
        }

        private void toolStripCBBatMod_DropDown(object sender, EventArgs e)
        {
            InhibitCB = false;
        }

        private void toolStripCBCustomers_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid)
                {
                    Bats.Tables[0].Rows[Bats.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }

        }

        private void toolStripCBBatMod_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid)
                {
                    Bats.Tables[0].Rows[Bats.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void toolStripCBCustomers_Enter(object sender, EventArgs e)
        {
            InhibitCB = false;
        }

        private void toolStripCBCustomers_Leave(object sender, EventArgs e)
        {
            InhibitCB = true;
        }

        private void toolStripCBBatMod_Enter(object sender, EventArgs e)
        {
            InhibitCB = false;
        }

        private void toolStripCBBatMod_Leave(object sender, EventArgs e)
        {
            InhibitCB = true;
        }

        private void toolStripCBSerNum_Enter(object sender, EventArgs e)
        {

        }
    }
}
