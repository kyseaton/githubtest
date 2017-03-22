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
using System.Threading;

namespace NewBTASProto
{

    public partial class frmVEWorkOrders : Form
    {
        // class wide variables
        DataSet WorkOrders = new DataSet();
        DataSet testList = new DataSet();
        int max;
        string BID;

        // current data
        string curTemp1;
        string curTemp2;
        string curTemp3;
        string curTemp4;
        string curTemp5;
        string curTemp6;
        string curTemp7;
        string curTemp8;
        string curTemp9;
        string curTemp10;
        string curTemp11;
        string curTemp12;

        // we use this bool to allow us to allow the databinding indext to be changed...
        bool Inhibit = true;
        public bool InhibitCB1 = true;         //workOrderStatusrCB
        bool InhibitCB2 = true;         //customerCB
        bool InhibitCB3 = true;         //serialNumCB
        public bool InhibitCB4 = true;         //workOrderCB

        public frmVEWorkOrders()
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

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT WorkOrderID,WorkOrderNumber,DateReceived,PlaneType,TailNumber,TestRequested,DateCompleted,OrderStatus,Notes,Batteries.BatteryModel,Batteries.BatterySerialNumber,Batteries.BatteryBCN,Batteries.CustomerName" +
                @" FROM WorkOrders LEFT JOIN Batteries ON WorkOrders.BID=Batteries.BID WHERE OrderStatus='Open' ORDER BY WorkOrders.WorkOrderNumber ASC";
            

            WorkOrders.Clear();
            OleDbConnection myAccessConn = null;
            // try to open the DB
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    myDataAdapter.Fill(WorkOrders);
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                
            }





            // Set the DataSource to the DataSet, and the DataMember
            bindingSource1.DataSource = null;
            bindingSource1.DataSource = WorkOrders;

            bindingSource1.DataMember = "Table";

            //work order stuff
            textBox1.DataBindings.Add("Text", bindingSource1, "WorkOrderNumber");
            dateTimePicker1.DataBindings.Add("Text", bindingSource1, "DateReceived");
            textBox3.DataBindings.Add("Text", bindingSource1, "PlaneType");
            textBox4.DataBindings.Add("Text", bindingSource1, "TailNumber");
            comboBox3.DataBindings.Add("Text", bindingSource1, "TestRequested");
            dateTimePicker2.DataBindings.Add("Text", bindingSource1, "DateCompleted");
            comboBox2.DataBindings.Add("Text", bindingSource1, "OrderStatus");

            //battery stuff
            textBox8.DataBindings.Add("Text", bindingSource1, "BatteryModel");
            comboBox1.DataBindings.Add("Text", bindingSource1, "BatterySerialNumber");
            textBox10.DataBindings.Add("Text", bindingSource1, "BatteryBCN");
            textBox11.DataBindings.Add("Text", bindingSource1, "CustomerName");
            textBox12.DataBindings.Add("Text", bindingSource1, "Notes");

            #endregion

            #region setup the combo boxes

            ComboBox WorkOrderCB = toolStripCBWorkOrders.ComboBox;
            WorkOrderCB.DisplayMember = "WorkOrderNumber";
            WorkOrderCB.DataSource = bindingSource1;


            //  Setup the drop down to contain all customers availible in the customer table

            // Open database containing all the customer names data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
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
                MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                
            }

            List<string> Customers = Custs.Tables[0].AsEnumerable().Select(x => x[0].ToString()).Distinct().ToList();
            Customers.Sort();
            Customers.Insert(0, "");
            ComboBox CustCB = toolStripCBCustomers.ComboBox;
            CustCB.DataSource = Customers;

 //           foreach (string x in Customers)
 //           {
 //               comboBox1.Items.Add(x);
 //           }

            //Now we'll set up the Battery Serial Number drop down, so the customer can re assign the battery associated with the work order

            // Open database containing all the customer names data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT BatterySerialNumber FROM Batteries ORDER BY BatterySerialNumber ASC";

            DataSet Serials = new DataSet();
            myAccessConn = null;
            // try to open the DB
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    myDataAdapter.Fill(Serials);
                    myAccessConn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                
            }

            List<string> SerialNums = Serials.Tables[0].AsEnumerable().Select(x => x[0].ToString()).Distinct().ToList();
            SerialNums.Sort();
            SerialNums.Insert(0, "");
            ComboBox SerCB = toolStripCBSerialNums.ComboBox;
            SerCB.DataSource = SerialNums;

            foreach (string x in SerialNums)
            {
                comboBox1.Items.Add(x);
            }

            #endregion


            //set the max

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT WorkOrderID FROM WorkOrders";

            DataSet countSet = new DataSet();
            myAccessConn = null;
            // try to open the DB
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    myDataAdapter.Fill(countSet);
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                
            }

            //if there are no work orders to load, then just set max to 1..
            if (WorkOrders.Tables[0].Rows.Count < 1)
            {
                max = 1;
            }
            else
            {
                max = countSet.Tables[0].AsEnumerable().Max(r => r.Field<int>("WorkOrderID"));
            }

            toolStripCBWorkOrderStatus.Text = "Open";
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
            //lastValid = false;
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // look up the information and change the items in associated textboxes
            if (comboBox1.Text == "")
            {
                return;
            }
            // Open database containing all the customer names data....
            string strAccessConn;
            string strAccessSelect;

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT * FROM Batteries Where BatterySerialNumber='" + comboBox1.Text + @"' ORDER BY BatterySerialNumber ASC";

            DataSet Serials = new DataSet();
            OleDbConnection myAccessConn = null;
            // try to open the DB
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    myDataAdapter.Fill(Serials);
                    myAccessConn.Close();                    
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                
            }

            textBox8.Text = Serials.Tables[0].Rows[0][1].ToString();
            textBox10.Text = Serials.Tables[0].Rows[0][3].ToString();
            textBox11.Text = Serials.Tables[0].Rows[0][5].ToString();
            BID = Serials.Tables[0].Rows[0][0].ToString();


        }

        private void loadTests()
        {

            testList.Clear();

            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT StepNumber,TestName,Notes FROM Tests WHERE WorkOrderNumber='" + toolStripCBWorkOrders.Text + @"' ORDER BY StepNumber ASC";

            
            OleDbConnection myAccessConn = null;
            // try to open the DB
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    myDataAdapter.Fill(testList);
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                
            }

            dataGridView1.DataSource = testList.Tables[0];
            dataGridView1.ClearSelection();

        }


        private void toolStripCBWorkOrders_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (InhibitCB4)
            {
                loadTests();
            }

            //Validate before moving
            else if (ValidateIt())
            {

                // move back
                InhibitCB4 = true;
                toolStripCBWorkOrders.SelectedIndex = bindingNavigator1.BindingSource.Position;
                InhibitCB4 = false;
            }
            else
            {
                loadTests();
                updateCurVals();
            }

        }

        private bool ValidateIt()
        {
            // do we need to validate?
            if (curTemp1 != textBox1.Text ||
                curTemp2 != dateTimePicker1.Text ||
                curTemp3 != textBox3.Text ||
                curTemp4 != textBox4.Text ||
                curTemp5 != comboBox3.Text ||
                curTemp6 != dateTimePicker2.Text ||
                curTemp7 != comboBox2.Text ||
                curTemp8 != textBox12.Text ||
                curTemp9 != comboBox1.Text ||
                curTemp10 != textBox8.Text ||
                curTemp11 != textBox10.Text ||
                curTemp12 != textBox11.Text)
            {
                // they don't match!
                // ask if the user is sure that they want to continue...
                DialogResult dialogResult = MessageBox.Show(this, "Looks like this record has been updated without being saved.  Are you sure you want to navigate away without saving?", "Click Yes to continue or No to return.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.No)
                {
                    updateCurVals();
                    lastValid = false;
                    return true;
                }
                else
                {
                    //sync everything..
                    updateCurVals();

                }
            }
            lastValid = true;
            return false;
        }

        private void UpdateView()
        {
            #region setup the binding

            //prevent lockups
            if (WorkOrders.Tables[0].Rows.Count == 0 && (toolStripCBCustomers.Text != "" || toolStripCBSerialNums.Text != "" || toolStripCBWorkOrderStatus.Text != "")) { ;}
            else if (toolStripCBWorkOrderStatus.Text == "") return;
            

            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";

            // show everything
            if (toolStripCBWorkOrderStatus.Text == "Open and Closed" && toolStripCBCustomers.Text == "" && toolStripCBSerialNums.Text == "")
            {
                strAccessSelect = @"SELECT WorkOrderID,WorkOrderNumber,DateReceived,PlaneType,TailNumber,TestRequested,DateCompleted,OrderStatus,Notes,Batteries.BatteryModel,Batteries.BatterySerialNumber,Batteries.BatteryBCN,Batteries.CustomerName" +
                @" FROM WorkOrders LEFT JOIN Batteries ON WorkOrders.BID=Batteries.BID WHERE OrderStatus <> 'Active' ORDER BY WorkOrders.WorkOrderNumber ASC";
            }
            else
            {
                strAccessSelect = @"SELECT WorkOrderID,WorkOrderNumber,DateReceived,PlaneType,TailNumber,TestRequested,DateCompleted,OrderStatus,Notes,Batteries.BatteryModel,Batteries.BatterySerialNumber,Batteries.BatteryBCN,Batteries.CustomerName" +
                @" FROM WorkOrders LEFT JOIN Batteries ON WorkOrders.BID=Batteries.BID WHERE " +
                (toolStripCBWorkOrderStatus.Text != "Open and Closed" ? ("OrderStatus='" + toolStripCBWorkOrderStatus.Text + "' ") : "OrderStatus <> 'Active' AND ") +
                (toolStripCBWorkOrderStatus.Text != "Open and Closed" && toolStripCBCustomers.Text != "" ? (" AND ") : " ") +
                (toolStripCBCustomers.Text != "" ? ("Batteries.CustomerName='" + toolStripCBCustomers.Text.Replace("'", "''") + "' ") : " ") +
                ((toolStripCBCustomers.Text != "" && toolStripCBSerialNums.Text != "") || (toolStripCBWorkOrderStatus.Text != "Open and Closed" && toolStripCBSerialNums.Text != "") ? (" AND ") : " ") +
                (toolStripCBSerialNums.Text != "" ? ("Batteries.BatterySerialNumber='" + toolStripCBSerialNums.Text.Replace("'", "''") + "' ") : " ") +
                @" ORDER BY WorkOrders.WorkOrderNumber ASC";
            }
           


            WorkOrders.Clear();
            OleDbConnection myAccessConn = null;
            // try to open the DB
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    myDataAdapter.Fill(WorkOrders);
                    myAccessConn.Close();
                }
                                    loadTests();
                if (comboBox2.Text != "Active" && toolStripCBWorkOrderStatus.Text != "Active") 
                { 
                    bindingNavigatorAddNewItem.Enabled = true; 
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            #endregion


        }

        int oldPositionWOS = 0;

        private void toolStripCBWorkOrderStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (InhibitCB1)
            {
                return;
            }

            if(bindingNavigatorAddNewItem.Enabled == false)
            {
                bindingNavigatorMovePreviousItem_Click(null,null);
            }

            //Validate before moving
            if (ValidateIt())
            {
                // move back
                InhibitCB1 = true;
                toolStripCBWorkOrderStatus.SelectedIndex = oldPositionWOS;
                InhibitCB1 = false;
            }
            else
            {
                oldPositionWOS = toolStripCBCustomers.SelectedIndex;
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid && toolStripCBWorkOrderStatus.Text != "Active" && WorkOrders.Tables[0].Rows.Count > 0)
                {
                    
                    WorkOrders.Tables[0].Rows[WorkOrders.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                    lastValid = false;
                }
                #region enable disable depending....

                if (toolStripCBWorkOrderStatus.Text == "Active")
                {
                    if (comboBox2.Items.Contains("Open")) { comboBox2.Items.Remove("Open"); }
                    if (comboBox2.Items.Contains("Closed")) { comboBox2.Items.Remove("Closed"); }
                    if (!comboBox2.Items.Contains("Active")) { comboBox2.Items.Add("Active"); }
                }
                else
                {
                    if (!comboBox2.Items.Contains("Open")) { comboBox2.Items.Add("Open"); }
                    if (!comboBox2.Items.Contains("Closed")) { comboBox2.Items.Add("Closed"); }
                    if (comboBox2.Items.Contains("Active")) { comboBox2.Items.Remove("Active"); }
                }

                if (toolStripCBWorkOrderStatus.Text == "Closed")
                {
                    //disable everything
                    textBox1.Enabled = false;
                    dateTimePicker1.Enabled = false;
                    textBox3.Enabled = false;
                    textBox4.Enabled = false;
                    comboBox3.Enabled = false;
                    dateTimePicker2.Enabled = false;
                    textBox12.Enabled = false;
                    comboBox1.Enabled = false;
                    textBox8.Enabled = false;
                    textBox10.Enabled = false;
                    textBox11.Enabled = false;
                    button1.Enabled = false;
                    comboBox2.Enabled = true;
                    saveToolStripButton.Enabled = true;
                    bindingNavigatorDeleteItem.Enabled = true;
                    bindingNavigatorAddNewItem.Enabled = true;

                }
                else if (toolStripCBWorkOrderStatus.Text == "Active")
                {
                    //disable everything
                    textBox1.Enabled = false;
                    dateTimePicker1.Enabled = false;
                    textBox3.Enabled = false;
                    textBox4.Enabled = false;
                    comboBox3.Enabled = false;
                    dateTimePicker2.Enabled = false;
                    textBox12.Enabled = false;
                    comboBox1.Enabled = false;
                    textBox8.Enabled = false;
                    textBox10.Enabled = false;
                    textBox11.Enabled = false;
                    button1.Enabled = false;
                    comboBox2.Enabled = false;
                    saveToolStripButton.Enabled = false;
                    bindingNavigatorDeleteItem.Enabled = false;
                    bindingNavigatorAddNewItem.Enabled = false;
                }
                else
                {
                    //enable everything
                    textBox1.Enabled = true;
                    dateTimePicker1.Enabled = true;
                    textBox3.Enabled = true;
                    textBox4.Enabled = true;
                    comboBox3.Enabled = true;
                    dateTimePicker2.Enabled = true;
                    textBox12.Enabled = true;
                    comboBox1.Enabled = true;
                    textBox8.Enabled = true;
                    textBox10.Enabled = true;
                    textBox11.Enabled = true;
                    button1.Enabled = true;
                    comboBox2.Enabled = true;
                    saveToolStripButton.Enabled = true;
                    bindingNavigatorDeleteItem.Enabled = true;
                    bindingNavigatorAddNewItem.Enabled = true;

                }

                #endregion
                UpdateView();
                updateCurVals();
            }


        }

        int oldPositionCusts = 0;

        private void toolStripCBCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (InhibitCB2)
            {
                return;
            }

            //Validate before moving
            if (ValidateIt())
            {
                // move back
                InhibitCB2 = true;
                toolStripCBCustomers.SelectedIndex = oldPositionCusts;
                InhibitCB2 = false;
            }
            else
            {
                oldPositionCusts = toolStripCBCustomers.SelectedIndex;
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid && toolStripCBWorkOrderStatus.Text != "Active" && WorkOrders.Tables[0].Rows.Count > 0)
                {
                    WorkOrders.Tables[0].Rows[WorkOrders.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                    lastValid = false;
                }
                UpdateView();
                updateCurVals();
            }

        }

        int oldPositionSerNums = 0;

        private void toolStripCBSerialNums_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (InhibitCB3)
            {
                return;
            }

            //Validate before moving
            if (ValidateIt())
            {
                // move back
                InhibitCB3 = true;
                toolStripCBSerialNums.SelectedIndex = oldPositionSerNums;
                InhibitCB3 = false;
            }
            else
            {
                oldPositionSerNums = toolStripCBSerialNums.SelectedIndex;
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid && toolStripCBWorkOrderStatus.Text != "Active" && WorkOrders.Tables[0].Rows.Count > 0)
                {
                    WorkOrders.Tables[0].Rows[WorkOrders.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                    lastValid = false;
                }
                UpdateView();
                updateCurVals();
            }

        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            if (bindingNavigator1.BindingSource.Position == -1)
            {
                string temp1 = textBox1.Text;
                string temp2 = dateTimePicker1.Text;
                string temp3 = textBox3.Text;
                string temp4 = textBox4.Text;
                string temp5 = comboBox3.Text;
                string temp6 = dateTimePicker2.Text;
                string temp7 = comboBox2.Text;
                string temp8 = textBox12.Text;
                string temp9 = comboBox1.Text;
                string temp10 = textBox8.Text;
                string temp11 = textBox10.Text;
                string temp12 = textBox11.Text;

                bindingNavigator1.BindingSource.AddNew();
                bindingNavigator1.BindingSource.Position = 0;

                textBox1.Text = temp1;
                dateTimePicker1.Text = temp2;
                textBox3.Text = temp3;
                textBox4.Text = temp4;
                comboBox3.Text = temp5;
                dateTimePicker2.Text = temp6;
                comboBox2.Text = temp7;
                textBox12.Text = temp8;
                comboBox1.Text = temp9;
                textBox8.Text = temp10;
                textBox10.Text = temp11;
                textBox11.Text = temp12;

                comboBox2.Text = "Open";
            }

            int origPos = bindingNavigator1.BindingSource.Position;

            if (!comboBox1.Items.Contains(comboBox1.Text))
            {
                MessageBox.Show(this, "The selected battery serial number is not in the database.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Inhibit = true;
                return;
            }
            else if (comboBox1.Text == "")
            {
                MessageBox.Show(this, "Please select a battery serial number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Inhibit = true;
                return;
            }

            else if (textBox1.Text.Contains(" "))
            {
                MessageBox.Show(this, "Work order names cannot have spaces in them.  Please correct and press save again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Inhibit = true;
                return;
            }
            else if (comboBox2.Text =="Open")
            {
                DataRowView current = (DataRowView)bindingSource1.Current;

                // we also need to check to see if the battery is already associated with an open order!
                string strAccessConn;
                string strAccessSelect;
                // Open database containing all the battery data....

                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";

                strAccessSelect = @"SELECT OrderStatus,Batteries.BatterySerialNumber" +
                    @" FROM WorkOrders LEFT JOIN Batteries ON WorkOrders.BID=Batteries.BID WHERE OrderStatus <> 'Closed' AND WorkOrderNumber <> '" + current["WorkOrderNumber"] + "'";

                DataSet Bats = new DataSet();
                OleDbConnection myAccessConn = null;
                // try to open the DB
                try
                {
                    myAccessConn = new OleDbConnection(strAccessConn);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Inhibit = true;
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
                        bindingNavigatorAddNewItem.Enabled = true; 
                        myAccessConn.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Inhibit = true;
                    return;
                }
                finally
                {
                   
                }

                DataRow[] foundRows = Bats.Tables[0].Select("BatterySerialNumber = '" + comboBox1.Text + "'");

                if (foundRows.Length != 0)
                {
                    Inhibit = true;
                    MessageBox.Show(this, "That battery is already assigned to an Open order", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            
            try
            {

                // set up the db Connection
                string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                OleDbConnection conn = new OleDbConnection(connectionString);
                

                //MAKE SURE YOU SELECT THE CURRENT ROW FOR DOUBLE SAVES!!!!!!!!!!!!!!!!!

                //get the current row
                DataRowView current = (DataRowView)bindingSource1.Current;

                // since this form displays and edits two forms, we will have to update both...
                // first test to see if the record already is in the database

                if (current["WorkOrderID"].ToString() != "")
                {
                    //record already exist as we need to do an update
                    //first update the WorkOrders table

                    string cmdStr = "UPDATE WorkOrders SET WorkOrderNumber='" + textBox1.Text.Replace("'", "''") +
                        "', DateReceived='" + dateTimePicker1.Text +
                        "', PlaneType='" + textBox3.Text.Replace("'", "''") +
                        "', TailNumber='" + textBox4.Text.Replace("'", "''") +
                        "', TestRequested='" + comboBox3.Text +
                        "', DateCompleted='" + dateTimePicker2.Text +
                        "', OrderStatus='" + comboBox2.Text +
                        "', Notes='" + textBox12.Text.Replace("'", "''") +
                        "', BatteryModel='" + textBox8.Text.Replace("'", "''") +
                        "', BatterySerialNumber='" + comboBox1.Text +
                        "', BatteryBCN='" + textBox10.Text.Replace("'", "''") +
                        "', CustomerName='" + textBox11.Text.Replace("'", "''") +
                        "', BID='" + BID +
                        "' WHERE WorkOrderID=" + current["WorkOrderID"].ToString();
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);


                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                    // Also update the workorder number in the other tables!
                    cmdStr = "UPDATE Tests SET WorkOrderNumber='" + textBox1.Text.Replace("'", "''") + "' WHERE WorkOrderNumber='" + current["WorkOrderNumber"].ToString() + "'";
                    cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                    cmdStr = "UPDATE ScanData SET BWO='" + textBox1.Text.Replace("'", "''") + "' WHERE BWO='" + current["WorkOrderNumber"].ToString() + "'";
                    cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }


                    //now update the combobox..
                    toolStripCBWorkOrders.ComboBox.Text = textBox1.Text.Replace("'", "''");
                    MessageBox.Show(this, textBox1.Text.Replace("'", "''") + " has been updated.", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                }
                else
                {
                    // we need to insert a new record...
                    // first check to see if the serial number is already in use.
                    string checkString = "SELECT * FROM WorkOrders WHERE WorkOrderNumber = '" + textBox1.Text.Replace("'", "''") + "'";
                    DataSet checkSet = new DataSet();
                    OleDbCommand checkCmd = new OleDbCommand(checkString, conn);
                    OleDbDataAdapter checkAdapter = new OleDbDataAdapter(checkCmd);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        checkAdapter.Fill(checkSet);
                        conn.Close();
                    }

                    if (checkSet.Tables[0].Rows.Count > 0)
                    {
                        //we already have that serial number in the DB
                        // tell the user about that and return...
                        MessageBox.Show(this, "That Work Order Number is already in the database!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        updateCurVals();
                        return;
                    }


                    // find the max value in the CustomerID column so we know what to assign to the new record
                    string cmdStr = "INSERT INTO WorkOrders (WorkOrderNumber, DateReceived, PlaneType, TailNumber, TestRequested, DateCompleted, OrderStatus, Notes, BatteryModel, BatterySerialNumber, BatteryBCN, CustomerName, BID) " +
                        "VALUES ('" +
                        textBox1.Text.Replace("'", "''") + "','" +
                        dateTimePicker1.Text + "','" +
                        textBox3.Text.Replace("'", "''") + "','" +
                        textBox4.Text.Replace("'", "''") + "','" +
                        comboBox3.Text.Replace("'", "''") + "','" +
                        dateTimePicker2.Text + "','" +
                        comboBox2.Text.Replace("'", "''") + "','" +                        
                        textBox12.Text.Replace("'", "''") + "','" +
                        textBox8.Text.Replace("'", "''") + "','" +
                        comboBox1.Text.Replace("'", "''") + "','" +
                        textBox10.Text.Replace("'", "''") + "','" +
                        textBox11.Text.Replace("'", "''") + "','" +
                        BID + "')";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);

                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                    // update the dataTable with the new customer ID also..
                    current[0] = max + 1;
                    max += 1;
                    MessageBox.Show(this, textBox1.Text + " has been created.", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                updateCurVals();

                //now we are going to save the notes on the test page...
                //first test to see if we have any tests before continuing
                if (testList.Tables[0].Rows.Count < 1) return;
                else
                {
                    dataGridView1.EndEdit();
                    for (int i = 0; i < testList.Tables[0].Rows.Count; i++ )
                    {
                        if (dataGridView1.Rows[i].Cells[2].Value.ToString().Replace("'", "''") != "")
                        {
                            string cmdStr = "UPDATE Tests SET Notes='" + dataGridView1.Rows[i].Cells[2].Value.ToString().Replace("'", "''") +
                                "' WHERE WorkOrderNumber='" + textBox1.Text.Replace("'", "''") + "' AND StepNumber='" + dataGridView1.Rows[i].Cells[0].Value.ToString() + "'";
                            OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                            lock (Main_Form.dataBaseLock)
                            {
                                conn.Open();
                                cmd.ExecuteNonQuery();
                                conn.Close();
                            }
                        }
                        else
                        {
                            string cmdStr = "UPDATE Tests SET Notes= Null WHERE WorkOrderNumber='" + textBox1.Text.Replace("'", "''") + "' AND StepNumber='" + dataGridView1.Rows[i].Cells[0].Value.ToString() + "'";
                            OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                            lock (Main_Form.dataBaseLock)
                            {
                                conn.Open();
                                cmd.ExecuteNonQuery();
                                conn.Close();
                            }
                        }

                    }
                }

                bindingNavigatorAddNewItem.Enabled = true;
                UpdateView();
                if (bindingSource1.Count > 1)
                {
                    bindingSource1.Position = origPos;
                }

                

            }// end try
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show(this, "Are you sure you want to remove this work order?", "Delete Record", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    // set up the db Connection
                    string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                    OleDbConnection conn = new OleDbConnection(connectionString);

                    //get the current row
                    DataRowView current = (DataRowView)bindingSource1.Current;

                    // first test to see if the record already is actually in the database

                    if (current["WorkOrderID"].ToString() != "")
                    {
                        // first delete the tests and scandata!
                        string cmdStr = "DELETE FROM Tests WHERE WorkOrderNumber='" + current["WorkOrderNumber"].ToString() + "'";
                        OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                        lock (Main_Form.dataBaseLock)
                        {
                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }

                        cmdStr = "DELETE FROM ScanData WHERE BWO='" + current["WorkOrderNumber"].ToString() + "'";
                        cmd = new OleDbCommand(cmdStr, conn);
                        lock (Main_Form.dataBaseLock)
                        {
                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }

                        cmdStr = "DELETE FROM WorkOrders WHERE WorkOrderID=" + current["WorkOrderID"].ToString();
                        cmd = new OleDbCommand(cmdStr, conn);
                        lock (Main_Form.dataBaseLock)
                        {
                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }

                        // Also update the binding source
                        WorkOrders.Tables[0].Rows[bindingNavigator1.BindingSource.Position].Delete();

                    }
                    else
                    {
                        MessageBox.Show(this, "That record was not in the DB. You must save it in order to delete it.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    UpdateView();
                    updateCurVals();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Deletion Error" + ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                //first see if there is anything to delete
                if (testList.Tables[0].Rows.Count < 1)
                {
                    MessageBox.Show(this, "No tests to delete!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                else if (MessageBox.Show(this, "Are you sure you want to remove this test?", "Delete Record", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    // set up the db Connection
                    string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                    OleDbConnection conn = new OleDbConnection(connectionString);

                    // first get rid of the scan data from the test
                    string cmdStr = "DELETE FROM ScanData WHERE BWO='" + textBox1.Text + "' AND STEP='" + testList.Tables[0].Rows[testList.Tables[0].Rows.Count - 1][0] + "'";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                    cmdStr = "DELETE FROM Tests WHERE WorkOrderNumber='" + textBox1.Text + "' AND StepNumber='" + testList.Tables[0].Rows[testList.Tables[0].Rows.Count - 1][0] + "'";
                    cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    // Also update the test datagrid view again
                    // use another thread to dealy the call
                    loadTests();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Deletion Error" + ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void frmVEWorkOrders_Load(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorMoveNextItem_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "Closed")
            {
                //disable everything
                textBox1.Enabled = false;
                dateTimePicker1.Enabled = false;
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                comboBox3.Enabled = false;
                dateTimePicker2.Enabled = false;
                textBox12.Enabled = false;
                comboBox1.Enabled = false;
                textBox8.Enabled = false;
                textBox10.Enabled = false;
                textBox11.Enabled = false;
                button1.Enabled = false;
                comboBox2.Enabled = true;
                saveToolStripButton.Enabled = true;
                bindingNavigatorDeleteItem.Enabled = true;
                bindingNavigatorAddNewItem.Enabled = true;
                button2.Visible = false;

            }
            else if (comboBox2.Text == "Active")
            {
                //disable everything
                textBox1.Enabled = false;
                dateTimePicker1.Enabled = false;
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                comboBox3.Enabled = false;
                dateTimePicker2.Enabled = false;
                textBox12.Enabled = false;
                comboBox1.Enabled = false;
                textBox8.Enabled = false;
                textBox10.Enabled = false;
                textBox11.Enabled = false;
                button1.Enabled = true;
                comboBox2.Enabled = false;
                saveToolStripButton.Enabled = false;
                bindingNavigatorDeleteItem.Enabled = false;
                bindingNavigatorAddNewItem.Enabled = false;
                button2.Visible = true;
            }
            else
            {
                //enable everything
                textBox1.Enabled = true;
                dateTimePicker1.Enabled = true;
                textBox3.Enabled = true;
                textBox4.Enabled = true;
                comboBox3.Enabled = true;
                dateTimePicker2.Enabled = true;
                textBox12.Enabled = true;
                comboBox1.Enabled = true;
                textBox8.Enabled = true;
                textBox10.Enabled = true;
                textBox11.Enabled = true;
                button1.Enabled = true;
                comboBox2.Enabled = true;
                saveToolStripButton.Enabled = true;
                bindingNavigatorDeleteItem.Enabled = true;
                bindingNavigatorAddNewItem.Enabled = true;
                button2.Visible = false;

            }

        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            Inhibit = true;
            if (toolStripCBWorkOrders.Text == "")
            {
                bindingNavigatorAddNewItem.Enabled = false;
                comboBox2.Text = "Open";
            }
            lastValid = false;
        }

        private void bindingNavigatorMovePreviousItem_Click(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid && toolStripCBWorkOrderStatus.Text != "Active" && WorkOrders.Tables[0].Rows.Count > 0)
                {
                    WorkOrders.Tables[0].Rows[WorkOrders.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                    lastValid = false;
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
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid && toolStripCBWorkOrderStatus.Text != "Active" && WorkOrders.Tables[0].Rows.Count > 0)
                {
                    WorkOrders.Tables[0].Rows[WorkOrders.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                    lastValid = false;
                }
            }
            catch
            {
                //do nothing
            }

        }

        private void toolStripCBWorkOrders_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid && toolStripCBWorkOrderStatus.Text != "Active" && WorkOrders.Tables[0].Rows.Count > 0)
                {
                    WorkOrders.Tables[0].Rows[WorkOrders.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                    lastValid = false;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void updateCurVals()
        {
            // update the current vars....
            //current data..
            curTemp1 = textBox1.Text;
            curTemp2 = dateTimePicker1.Text;
            curTemp3 = textBox3.Text;
            curTemp4 = textBox4.Text;
            curTemp5 = comboBox3.Text;
            curTemp6 = dateTimePicker2.Text;
            curTemp7 = comboBox2.Text;
            curTemp8 = textBox12.Text;
            curTemp9 = comboBox1.Text;
            curTemp10 = textBox8.Text;
            curTemp11 = textBox10.Text;
            curTemp12 = textBox11.Text;
        }

        bool lastValid = false;

        private void bindingNavigator1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid && toolStripCBWorkOrderStatus.Text != "Active" && WorkOrders.Tables[0].Rows.Count > 0)
                {
                    WorkOrders.Tables[0].Rows[WorkOrders.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }
            
        }

        private void frmVEWorkOrders_FormClosing(object sender, FormClosingEventArgs e)
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

        private void bindingNavigator1_Validating(object sender, CancelEventArgs e)
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
            }
        }

        private void toolStripCBWorkOrderStatus_Enter(object sender, EventArgs e)
        {
            InhibitCB1 = false;
        }

        private void toolStripCBWorkOrderStatus_Leave(object sender, EventArgs e)
        {
            InhibitCB1 = true;
        }

        private void toolStripCBCustomers_Enter(object sender, EventArgs e)
        {
            InhibitCB2 = false;
        }

        private void toolStripCBCustomers_Leave(object sender, EventArgs e)
        {
            InhibitCB2 = true;
        }

        private void toolStripCBSerialNums_Enter(object sender, EventArgs e)
        {
            InhibitCB3 = false;
        }

        private void toolStripCBSerialNums_Leave(object sender, EventArgs e)
        {
            InhibitCB3 = true;
        }

        private void toolStripCBWorkOrders_Enter(object sender, EventArgs e)
        {
            InhibitCB4 = false;
        }

        private void toolStripCBWorkOrders_Leave(object sender, EventArgs e)
        {
            InhibitCB4 = true;
        }

        private void bindingNavigator1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            Inhibit = false;
            bindingNavigator1.Focus();
        }

        private void toolStripCBSerialNums_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid && toolStripCBWorkOrderStatus.Text != "Active" && WorkOrders.Tables[0].Rows.Count > 0)
                {
                    WorkOrders.Tables[0].Rows[WorkOrders.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                    lastValid = false;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void toolStripCBCustomers_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid && toolStripCBWorkOrderStatus.Text != "Active" && WorkOrders.Tables[0].Rows.Count > 0)
                {
                    WorkOrders.Tables[0].Rows[WorkOrders.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                    lastValid = false;
                }
            }
            catch
            {
                //do nothing
            }
        }


        private void toolStripCBWorkOrderStatus_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && lastValid && toolStripCBWorkOrderStatus.Text != "Active" && WorkOrders.Tables[0].Rows.Count > 0)
                {
                    //WorkOrders.Tables[0].Rows[WorkOrders.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                    lastValid = false;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void frmVEWorkOrders_Shown(object sender, EventArgs e)
        {
            //bindingNavigatorAddNewItem.PerformClick();
            #region enable disable depending....

            if (toolStripCBWorkOrderStatus.Text == "Active")
            {
                if (comboBox2.Items.Contains("Open")) { comboBox2.Items.Remove("Open"); }
                if (comboBox2.Items.Contains("Closed")) { comboBox2.Items.Remove("Closed"); }
                if (!comboBox2.Items.Contains("Active")) { comboBox2.Items.Add("Active"); }
            }
            else
            {
                if (!comboBox2.Items.Contains("Open")) { comboBox2.Items.Add("Open"); }
                if (!comboBox2.Items.Contains("Closed")) { comboBox2.Items.Add("Closed"); }
                if (comboBox2.Items.Contains("Active")) { comboBox2.Items.Remove("Active"); }
            }

            if (toolStripCBWorkOrderStatus.Text == "Closed")
            {
                //disable everything
                textBox1.Enabled = false;
                dateTimePicker1.Enabled = false;
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                comboBox3.Enabled = false;
                dateTimePicker2.Enabled = false;
                textBox12.Enabled = false;
                comboBox1.Enabled = false;
                textBox8.Enabled = false;
                textBox10.Enabled = false;
                textBox11.Enabled = false;
                button1.Enabled = false;
                comboBox2.Enabled = true;
                saveToolStripButton.Enabled = true;
                bindingNavigatorDeleteItem.Enabled = true;
                bindingNavigatorAddNewItem.Enabled = true;

            }
            else if (toolStripCBWorkOrderStatus.Text == "Active")
            {
                //disable everything
                textBox1.Enabled = false;
                dateTimePicker1.Enabled = false;
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                comboBox3.Enabled = false;
                dateTimePicker2.Enabled = false;
                textBox12.Enabled = false;
                comboBox1.Enabled = false;
                textBox8.Enabled = false;
                textBox10.Enabled = false;
                textBox11.Enabled = false;
                button1.Enabled = false;
                comboBox2.Enabled = false;
                saveToolStripButton.Enabled = false;
                bindingNavigatorDeleteItem.Enabled = false;
                bindingNavigatorAddNewItem.Enabled = false;
            }
            else
            {
                //enable everything
                textBox1.Enabled = true;
                dateTimePicker1.Enabled = true;
                textBox3.Enabled = true;
                textBox4.Enabled = true;
                comboBox3.Enabled = true;
                dateTimePicker2.Enabled = true;
                textBox12.Enabled = true;
                comboBox1.Enabled = true;
                textBox8.Enabled = true;
                textBox10.Enabled = true;
                textBox11.Enabled = true;
                button1.Enabled = true;
                comboBox2.Enabled = true;
                saveToolStripButton.Enabled = true;
                bindingNavigatorDeleteItem.Enabled = true;
                bindingNavigatorAddNewItem.Enabled = true;

            }

            #endregion
            UpdateView();
            updateCurVals();

            bindingNavigatorAddNewItem.PerformClick();
        }

        private void comboBox1_Enter(object sender, EventArgs e)
        {
            Inhibit = true;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if ((this.ActiveControl == toolStripCBWorkOrderStatus.ComboBox ) && (keyData == Keys.Return))
            {
                if (ValidateIt()) { return true; }
                InhibitCB1 = false;
                toolStripCBWorkOrders_TextChanged(null, null);
                UpdateView();
                updateCurVals();
                return true;
            }
            else if((this.ActiveControl == toolStripCBSerialNums.ComboBox) && (keyData == Keys.Return))
            {
                if (ValidateIt()) { return true; }
                InhibitCB3 = false;
                toolStripCBWorkOrders_TextChanged(null, null);
                UpdateView();
                updateCurVals();
                return true;
            }
            else if((this.ActiveControl == toolStripCBCustomers.ComboBox) && (keyData == Keys.Return))
            {
                if (ValidateIt()) { return true; }
                InhibitCB2 = false;
                toolStripCBWorkOrders_TextChanged(null, null);
                UpdateView();
                updateCurVals();
                return true;
            }
            else
            {
                return base.ProcessCmdKey(ref msg, keyData);
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                ThreadPool.QueueUserWorkItem(s =>
                {
                    int i = 0;

                    // we double clicked on the first column do we have a Graph or test report windo open?
                    FormCollection fc = Application.OpenForms;

                    // first the graphics_form section
                    foreach (Form frm in fc)
                    {
                        if (frm is Graphics_Form)
                        {
                            if (frm.WindowState == FormWindowState.Minimized)
                            {
                                frm.WindowState = FormWindowState.Normal;
                            }

                            Graphics_Form to_control = (Graphics_Form)frm;
                            this.Invoke((MethodInvoker)delegate()
                            {
                                to_control.comboBox1.SelectedValue = toolStripCBWorkOrders.Text;
                                to_control.comboBox3.SelectedValue = toolStripCBWorkOrders.Text;
                            });

                            while (to_control.comboBox2.Items.Count < 1)
                            {
                                Thread.Sleep(100);
                                if (i > 20)
                                {
                                    // we timed out so we need to return...
                                    return;
                                }
                            }

                            i = 0;
                            while (to_control.comboBox4.Items.Count < 1)
                            {
                                Thread.Sleep(100);
                                if (i > 20)
                                {
                                    // we timed out so we need to return...
                                    return;
                                }
                            }

                            Thread.Sleep(200);

                            this.Invoke((MethodInvoker)delegate()
                            {
                                to_control.comboBox2.SelectedIndex = e.RowIndex + 1;
                                to_control.comboBox4.SelectedIndex = e.RowIndex + 1;
                            });

                        }

                    }

                    // and now the reports_form section
                    foreach (Form frm in fc)
                    {
                        if (frm is Reports_Form)
                        {
                            if (frm.WindowState == FormWindowState.Minimized)
                            {
                                frm.WindowState = FormWindowState.Normal;
                            }

                            Reports_Form to_control = (Reports_Form)frm;
                            this.Invoke((MethodInvoker)delegate()
                            {
                                to_control.comboBox1.SelectedValue = toolStripCBWorkOrders.Text;
                            });

                            i = 0;
                            while (to_control.comboBox2.Items.Count < 1)
                            {
                                Thread.Sleep(100);
                                if (i > 20)
                                {
                                    // we timed out so we need to return...
                                    return;
                                }
                            }

                            Thread.Sleep(200);

                            this.Invoke((MethodInvoker)delegate()
                            {
                                if (to_control.comboBox2.GetItemText(to_control.comboBox2.Items[1]).Contains("Water"))
                                {
                                    to_control.comboBox2.SelectedIndex = e.RowIndex + 2;
                                }
                                else
                                {
                                    to_control.comboBox2.SelectedIndex = e.RowIndex + 1;
                                }
                                
                            });

                            Thread.Sleep(200);

                            this.Invoke((MethodInvoker)delegate()
                            {
                                to_control.comboBox3.SelectedIndex = 1;
                            });

                            return;
                        }
                    }
                });


            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //now we are going to save the notes on the test page...
            //first test to see if we have any tests before continuing
            try
            {
                // set up the db Connection
                string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                OleDbConnection conn = new OleDbConnection(connectionString);
                
                if (testList.Tables[0].Rows.Count < 1) return;
                else
                {
                    dataGridView1.EndEdit();
                    for (int i = 0; i < testList.Tables[0].Rows.Count; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[2].Value.ToString().Replace("'", "''") != "")
                        {
                            string cmdStr = "UPDATE Tests SET Notes='" + dataGridView1.Rows[i].Cells[2].Value.ToString().Replace("'", "''") +
                                "' WHERE WorkOrderNumber='" + textBox1.Text.Replace("'", "''") + "' AND StepNumber='" + dataGridView1.Rows[i].Cells[0].Value.ToString() + "'";
                            OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                            lock (Main_Form.dataBaseLock)
                            {
                                conn.Open();
                                cmd.ExecuteNonQuery();
                                conn.Close();
                            }
                        }
                        else
                        {
                            string cmdStr = "UPDATE Tests SET Notes= Null WHERE WorkOrderNumber='" + textBox1.Text.Replace("'", "''") + "' AND StepNumber='" + dataGridView1.Rows[i].Cells[0].Value.ToString() + "'";
                            OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                            lock (Main_Form.dataBaseLock)
                            {
                                conn.Open();
                                cmd.ExecuteNonQuery();
                                conn.Close();
                            }
                        }

                    }
                }

                MessageBox.Show(this, "Notes Saved","Success!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch(Exception ex)
            {
                MessageBox.Show(this, "Note Saving Error" + ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
