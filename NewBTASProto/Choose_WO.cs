using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace NewBTASProto
{
    public partial class Choose_WO : Form
    {
        int selectedChannel;
        string[] oldSelect;

        public Choose_WO(int channel, string fromGrid)
        {
            // save the channel you are working on
            selectedChannel = channel;

            //split up the instring so you can highlight the previously selected items
            char[] delims = { ' ' };
            oldSelect = fromGrid.Split(delims);
            
            //  now onto the form stuff...
            InitializeComponent();
            loadWorkOrderLists();
        }

        private void loadWorkOrderLists()
        {

            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT WorkOrderNumber,CustomerName,DateReceived FROM WorkOrders WHERE OrderStatus='Open'";

            DataSet workOrderList1 = new DataSet();
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

                lock (Main_Form.dataBaseLock)
                {
                    myAccessConn.Open();
                    myDataAdapter.Fill(workOrderList1, "ScanData");
                    myAccessConn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                return;
            }
            finally
            {

            }

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            DataSet workOrderList2;
            DataRow tempRow;

            foreach (string oldWO in oldSelect)
            {
                if (oldWO == "") { ;}  // do nothing
                else
                {
                    strAccessSelect = @"SELECT WorkOrderNumber,CustomerName,DateReceived FROM WorkOrders WHERE WorkOrderNumber='" + oldWO + "'";
                    // Add aditional rows to show the currently selected WOs at top of workOrderList1.Tables["ScanData"]

                    workOrderList2 = new DataSet();
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

                        lock (Main_Form.dataBaseLock)
                        {
                            myAccessConn.Open();
                            myDataAdapter.Fill(workOrderList2, "ScanData");
                            myAccessConn.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                        return;
                    }

                    try
                    {
                        //Now we need to add the records in!  ,,
                        tempRow = workOrderList1.Tables["ScanData"].NewRow();
                        tempRow["WorkOrderNumber"] = workOrderList2.Tables["ScanData"].Rows[0][0];
                        tempRow["CustomerName"] = workOrderList2.Tables["ScanData"].Rows[0][1];
                        tempRow["DateReceived"] = workOrderList2.Tables["ScanData"].Rows[0][2];

                        workOrderList1.Tables["ScanData"].Rows.InsertAt(tempRow, 0);
                    }
                    catch
                    {
                        // do nothing for now...
                    }
                }
            }
              

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            dataGridView1.DataSource = workOrderList1.Tables["ScanData"];

        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            string temp = "";
            int count = 0;

            //Update the DB to show that the old Work Orders are now Open
            // set up the db Connection
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            OleDbConnection conn = new OleDbConnection(connectionString);

            string cmdStr = "";
            OleDbCommand cmd;

            foreach (string oldWO in oldSelect)
            {
                cmdStr = "UPDATE WorkOrders SET OrderStatus='Open' WHERE WorkOrderNumber='" + oldWO + "'";
                cmd = new OleDbCommand(cmdStr, conn);
                lock (Main_Form.dataBaseLock)
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }

            //Update the DB to show that the new Work Orders are now active
            // set up the db Connection            
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                count++;
                if (count > 3)
                {
                    MessageBox.Show("Maximum of 3 Work Orders Per Channel!");
                    break;
                }
                temp = dataGridView1[0, row.Index].Value + " " + temp;
                cmdStr = "UPDATE WorkOrders SET OrderStatus='Active' WHERE WorkOrderNumber='" + dataGridView1[0, row.Index].Value + "'";
                cmd = new OleDbCommand(cmdStr, conn);
                lock (Main_Form.dataBaseLock)
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }


            ((Main_Form)this.Owner).updateWOC(selectedChannel,temp);
            this.Dispose();
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            //Update the DB to show that the old Work Orders are now Open
            // set up the db Connection
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            OleDbConnection conn = new OleDbConnection(connectionString);

            string cmdStr = "";
            OleDbCommand cmd;

            foreach (string oldWO in oldSelect)
            {
                cmdStr = "UPDATE WorkOrders SET OrderStatus='Open' WHERE WorkOrderNumber='" + oldWO + "'";
                cmd = new OleDbCommand(cmdStr, conn);
                lock (Main_Form.dataBaseLock)
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }


            //listBox1.ClearSelected();
            string temp = "";

            ((Main_Form)this.Owner).updateWOC(selectedChannel, temp);
            this.Dispose();
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            string temp = "";
            int count = 0;

            //Update the DB to show that the old Work Orders are now Open
            // set up the db Connection
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            OleDbConnection conn = new OleDbConnection(connectionString);

            string cmdStr = "";
            OleDbCommand cmd;

            foreach (string oldWO in oldSelect)
            {
                cmdStr = "UPDATE WorkOrders SET OrderStatus='Open' WHERE WorkOrderNumber='" + oldWO + "'";
                cmd = new OleDbCommand(cmdStr, conn);
                lock (Main_Form.dataBaseLock)
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }

            //Update the DB to show that the new Work Orders are now active
            // set up the db Connection            
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                count++;
                if (count > 3)
                {
                    MessageBox.Show("Maximum of 3 Work Orders Per Channel!");
                    break;
                }
                temp += dataGridView1[0, row.Index].Value + " ";
                cmdStr = "UPDATE WorkOrders SET OrderStatus='Active' WHERE WorkOrderNumber='" + dataGridView1[0, row.Index].Value + "'";
                cmd = new OleDbCommand(cmdStr, conn);
                lock (Main_Form.dataBaseLock)
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }


            ((Main_Form)this.Owner).updateWOC(selectedChannel, temp);
            this.Dispose();
        }

    }
}
