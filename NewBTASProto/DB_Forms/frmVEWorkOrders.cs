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

        public frmVEWorkOrders()
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

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\DB\BTS16NV.MDB";
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
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(WorkOrders);

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

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\DB\BTS16NV.MDB";
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

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\DB\BTS16NV.MDB";
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
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(Serials);

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

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\DB\BTS16NV.MDB";
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
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(countSet);

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

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\DB\BTS16NV.MDB";
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
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(Serials);

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

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT StepNumber,TestName,Notes FROM Tests WHERE WorkOrderNumber='" + toolStripCBWorkOrders.Text + @"' ORDER BY StepNumber ASC";

            
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
                myDataAdapter.Fill(testList);

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

            dataGridView1.DataSource = testList.Tables[0];
            dataGridView1.ClearSelection();

        }

        private void toolStripCBWorkOrders_SelectedIndexChanged(object sender, EventArgs e)
        {
            loadTests();
        }

        private void UpdateView()
        {
            #region setup the binding

            //prevent lockups
            if (WorkOrders.Tables[0].Rows.Count == 0 && (toolStripCBCustomers.Text != "" || toolStripCBSerialNums.Text != "" || toolStripCBWorkOrderStatus.Text != "")) { ;}
            else if (toolStripCBWorkOrders.Text == "" || toolStripCBWorkOrderStatus.Text == "") return;
            

            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\DB\BTS16NV.MDB";

            // show everything
            if ( toolStripCBWorkOrderStatus.Text == "All" && toolStripCBCustomers.Text == "" && toolStripCBSerialNums.Text == "")
            {
                strAccessSelect = @"SELECT WorkOrderID,WorkOrderNumber,DateReceived,PlaneType,TailNumber,TestRequested,DateCompleted,OrderStatus,Notes,Batteries.BatteryModel,Batteries.BatterySerialNumber,Batteries.BatteryBCN,Batteries.CustomerName" +
                @" FROM WorkOrders LEFT JOIN Batteries ON WorkOrders.BID=Batteries.BID ORDER BY WorkOrders.WorkOrderNumber ASC";
            }
            else
            {
                strAccessSelect = @"SELECT WorkOrderID,WorkOrderNumber,DateReceived,PlaneType,TailNumber,TestRequested,DateCompleted,OrderStatus,Notes,Batteries.BatteryModel,Batteries.BatterySerialNumber,Batteries.BatteryBCN,Batteries.CustomerName" +
                @" FROM WorkOrders LEFT JOIN Batteries ON WorkOrders.BID=Batteries.BID WHERE " + 
                (toolStripCBWorkOrderStatus.Text != "All" ? ("OrderStatus='" + toolStripCBWorkOrderStatus.Text + "' " ) : " ") +
                (toolStripCBWorkOrderStatus.Text != "All" && toolStripCBCustomers.Text != "" ? (" AND ") : " ") +
                (toolStripCBCustomers.Text != "" ? ("Batteries.CustomerName='" + toolStripCBCustomers.Text.Replace("'", "''") + "' ") : " ") +
                ((toolStripCBCustomers.Text != "" && toolStripCBSerialNums.Text != "") || (toolStripCBWorkOrderStatus.Text != "All" && toolStripCBSerialNums.Text != "") ? (" AND ") : " ") +
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
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(WorkOrders);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                return;
            }
            finally
            {
                myAccessConn.Close();
                loadTests();
            }

            #endregion


        }

        private void toolStripCBWorkOrderStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateView();
        }

        private void toolStripCBCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateView();
        }

        private void toolStripCBSerialNums_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateView();
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {

                // set up the db Connection
                string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\DB\BTS16NV.MDB";
                OleDbConnection conn = new OleDbConnection(connectionString);
                conn.Open();

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
                    cmd.ExecuteNonQuery();

                }
                else
                {
                    // we need to insert a new record...
                    // find the max value in the CustomerID column so we know what to assign to the new record
                    string cmdStr = "INSERT INTO WorkOrders (WorkOrderID, WorkOrderNumber, DateReceived, PlaneType, TailNumber, TestRequested, DateCompleted, OrderStatus, Notes, BatteryModel, BatterySerialNumber, BatteryBCN, CustomerName, BID) " +
                        "VALUES (" + (max + 1).ToString() + ",'" +
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
                    cmd.ExecuteNonQuery();

                    // update the dataTable with the new customer ID also..
                    current[0] = max + 1;
                    max += 1;


                }

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
                            cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            string cmdStr = "UPDATE Tests SET Notes= Null WHERE WorkOrderNumber='" + textBox1.Text.Replace("'", "''") + "' AND StepNumber='" + dataGridView1.Rows[i].Cells[0].Value.ToString() + "'";
                            OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                            cmd.ExecuteNonQuery();
                        }

                    }
                }
                conn.Close();

            }// end try
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Are you sure you want to remove this work order?", "Delete Record", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    // set up the db Connection
                    string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\DB\BTS16NV.MDB";
                    OleDbConnection conn = new OleDbConnection(connectionString);
                    conn.Open();

                    //get the current row
                    DataRowView current = (DataRowView)bindingSource1.Current;

                    // first test to see if the record already is actually in the database

                    if (current["WorkOrderID"].ToString() != "")
                    {
                        //record already exist as we need to do an update

                        // first delete the tests and scandata!
                        string cmdStr = "DELETE FROM Tests WHERE WorkOrderNumber='" + current["WorkOrderNumber"].ToString() + "'";
                        OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                        cmd.ExecuteNonQuery();

                        cmdStr = "DELETE FROM ScanData WHERE BWO='" + current["WorkOrderNumber"].ToString() + "'";
                        cmd = new OleDbCommand(cmdStr, conn);
                        cmd.ExecuteNonQuery();

                        cmdStr = "DELETE FROM WorkOrders WHERE WorkOrderID=" + current["WorkOrderID"].ToString();
                        cmd = new OleDbCommand(cmdStr, conn);
                        cmd.ExecuteNonQuery();

                        // Also update the binding source
                        WorkOrders.Tables[0].Rows[bindingNavigator1.BindingSource.Position].Delete();

                    }
                    else
                    {
                        MessageBox.Show("That record was not in the DB. You must save it in order to delete it.");
                    }
                    conn.Close();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Deletion Error" + ex.ToString());
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                //first see if there is anything to delete
                if (testList.Tables[0].Rows.Count < 1)
                {
                    MessageBox.Show("No tests to delete!");
                }

                else if (MessageBox.Show("Are you sure you want to remove this test?", "Delete Record", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    // set up the db Connection
                    string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\DB\BTS16NV.MDB";
                    OleDbConnection conn = new OleDbConnection(connectionString);
                    conn.Open();

                    // first get rid of the scan data from the test
                    string cmdStr = "DELETE FROM ScanData WHERE BWO='" + textBox1.Text + "' AND STEP='" + testList.Tables[0].Rows[testList.Tables[0].Rows.Count - 1][0] + "'";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    cmd.ExecuteNonQuery();

                    cmdStr = "DELETE FROM Tests WHERE WorkOrderNumber='" + textBox1.Text + "' AND StepNumber='" + testList.Tables[0].Rows[testList.Tables[0].Rows.Count - 1][0] + "'";
                    cmd = new OleDbCommand(cmdStr, conn);
                    cmd.ExecuteNonQuery();
                    // Also update the test datagrid view again
                    // use another thread to dealy the call
                    conn.Close();
                    loadTests();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Deletion Error" + ex.ToString());
            }

        }

        private void frmVEWorkOrders_Load(object sender, EventArgs e)
        {

        }

    }
}
