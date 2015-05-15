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

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
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

                myAccessConn.Open();
                myDataAdapter.Fill(workOrderList1, "ScanData");

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

            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                count++;
                if (count > 3)
                {
                    MessageBox.Show("Maximum of 3 Work Orders Per Channel!");
                        break;
                }
                temp += dataGridView1[0, row.Index].Value + " ";
 
            }
            

            ((Main_Form)this.Owner).updateWOC(selectedChannel,temp);
            this.Dispose();
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            //listBox1.ClearSelected();
            string temp = "";

            ((Main_Form)this.Owner).updateWOC(selectedChannel, temp);
            this.Dispose();
        }

    }
}
