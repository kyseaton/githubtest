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
using System.IO.Ports;
using System.Threading;


namespace NewBTASProto
{
    public partial class Main_Form : Form
    {
        public Main_Form()
        {
            try
            {
                SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
                InitializeComponent();
                Initialize_Menus_Tools();
                Initialize_Operators_CB();

                InitializeGrid();
                InitializeTimers();
                Scan();
                

            }
            catch(Exception ex)
            {
                MessageBox.Show("In Main_Form:  " + ex.ToString());
            }


        }

        public void Initialize_Operators_CB()
        {
            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
            strAccessSelect = @"SELECT * FROM Operators";
            OleDbConnection myAccessConn;
            DataSet operators = new DataSet();

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
                myDataAdapter.Fill(operators, "Operators");

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

            this.comboBox1.DisplayMember = "OperatorName";
            this.comboBox1.ValueMember = "OperatorName";
            this.comboBox1.DataSource = operators.Tables["Operators"];


        }

        /// <summary>
        /// This function will update all of the Menus with the appropriate values
        /// </summary>
        private void Initialize_Menus_Tools()
        {
            if (GlobalVars.useF) { this.fahrenheitToolStripMenuItem.Checked = true; }
            else{this.centigradeToolStripMenuItem.Checked = true;}

            if (GlobalVars.Pos2Neg) { this.positiveToNegativeToolStripMenuItem.Checked = true; }
            else { this.negativeToPositiveToolStripMenuItem.Checked = true; }

            toolStripStatusLabel1.Text = "Version:  " + Application.ProductVersion;

            label10.Text = GlobalVars.businessName;


        }



        private void toolStripStatusLabel3_Click(object sender, EventArgs e)
        {

        }

        private void test3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Graphics_Form gf = new Graphics_Form();
            gf.Show();
   
        }

        private void testToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Reports_Form rf = new Reports_Form();
            rf.Show();
        }

        private void Main_Form_FormClosed(object sender, FormClosedEventArgs e)
        {
            SaveGlobals();
            Application.Exit();
        }

        private void SaveGlobals()
        {

            string strAccessConn;
            string strUpdateCMD;
            // Open database containing all the battery data....
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
            strUpdateCMD = "UPDATE Options SET Degree='" + (GlobalVars.useF ? "F." : "C.") + "', CellOrder='" + (GlobalVars.Pos2Neg ? "Pos. to Neg." : "Neg. to Pos.") + "', BusinessName='"+ GlobalVars.businessName+"';";
            OleDbConnection myAccessConn;

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
                OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                myAccessConn.Open();
                myAccessCommand.ExecuteNonQuery();

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

        }

        private void bussinessNameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Business_Name bn = new Business_Name();
            bn.Owner = this;
            bn.Show();
        }

        private void centigradeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.centigradeToolStripMenuItem.Checked == false) 
            { 
                this.centigradeToolStripMenuItem.Checked = true;
                this.fahrenheitToolStripMenuItem.Checked = false;
                GlobalVars.useF = false;
            }
        }

        private void fahrenheitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.fahrenheitToolStripMenuItem.Checked == false)
            {
                this.fahrenheitToolStripMenuItem.Checked = true;
                this.centigradeToolStripMenuItem.Checked = false;
                GlobalVars.useF = true;
            }

        }

        private void negativeToPositiveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.negativeToPositiveToolStripMenuItem.Checked == false)
            {
                this.negativeToPositiveToolStripMenuItem.Checked = true;
                this.positiveToNegativeToolStripMenuItem.Checked = false;
                GlobalVars.Pos2Neg = false;
            }
        }

        private void positiveToNegativeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.positiveToNegativeToolStripMenuItem.Checked == false)
            {
                this.positiveToNegativeToolStripMenuItem.Checked = true;
                this.negativeToPositiveToolStripMenuItem.Checked = false;
                GlobalVars.Pos2Neg = true;
            }
        }

        private void newCustomBatteryToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        public void updateBusiness() 
        {
            label10.Text = GlobalVars.businessName;
        }

        internal void updateWOC(int channel, string workOrder)
        {
            d.Rows[channel][1] = workOrder;
        }


        private void btnGetSerialPorts_Click_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }


        private void button1_Click(object sender, EventArgs e)
        {



            // create a thread to go through and look for the stations, this way the UI will still work while the search is happening
            ThreadPool.QueueUserWorkItem(s =>
            {
                this.Invoke((MethodInvoker)delegate
                    {
                        // start by disabling the button while we look for stations
                        button1.Enabled = false;
                        // also disable the grid, so the user cannot interfere with the search
                        dataGridView1.Enabled = false;
                        //select the first row as your selected cell
                        dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[0];
                        dataGridView1.ClearSelection();
                    });
                Thread.Sleep(500);

                // turn on all of the in use buttons
                for (int i = 0; i < 16; i++)
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        d.Rows[i][4] = true;
                        dataGridView1.Rows[i].Cells[4].Style.BackColor = Color.Red;
                    });
                }
                // here is the for loop we'll use to look for cscans
                for (int i = 0; i < 15; i++)
                {

                    this.Invoke((MethodInvoker)delegate
                    {
                        dataGridView1.CurrentCell = dataGridView1.Rows[i].Cells[0];
                        dataGridView1.ClearSelection();
                    });

                    //give it time to check the channel
                    Thread.Sleep(900);

                    // move the current channel
                    this.Invoke((MethodInvoker)delegate
                    {
                        dataGridView1.CurrentCell = dataGridView1.Rows[i+1].Cells[0];
                        dataGridView1.ClearSelection();
                    });

                    // wait again
                    Thread.Sleep(100);
                    this.Invoke((MethodInvoker)delegate
                    {
                        if (dataGridView1.Rows[i].Cells[4].Style.BackColor == Color.Red)
                        {
                            d.Rows[i][4] = false;
                            dataGridView1.Rows[i].Cells[4].Style.BackColor = Color.Gainsboro;
                        }
                    });
                }

                //Finally take care of the last channel
                //give it time to check the channel
                Thread.Sleep(900);

                // move back to channel 0
                this.Invoke((MethodInvoker)delegate
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[0];
                    dataGridView1.ClearSelection();
                });

                // wait again
                Thread.Sleep(100);
                this.Invoke((MethodInvoker)delegate
                {
                    if (dataGridView1.Rows[15].Cells[4].Style.BackColor == Color.Red)
                    {
                        d.Rows[15][4] = false;
                        dataGridView1.Rows[15].Cells[4].Style.BackColor = Color.Gainsboro;
                    }
                });

                //reenable the button before exit
                this.Invoke((MethodInvoker)delegate
                {
                    // start by disabling the button while we look for stations
                    button1.Enabled = true;
                    // also disable the grid, so the user cannot interfere with the search
                    dataGridView1.Enabled = true;
                });

            });                     // end thread



        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void Main_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            // tell those threadpool work items to stop!!!!!
            cPollIC.Cancel();
            cPollCScans.Cancel();
            sequentialScanT.Cancel();
            // make sure it takes...
            Thread.Sleep(500);      
        }

        private void highlightCurrentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (highlightCurrentToolStripMenuItem.Checked == false)
            {
                highlightCurrentToolStripMenuItem.Checked = true;
                GlobalVars.highlightCurrent = true;
            }
            else
            {
                highlightCurrentToolStripMenuItem.Checked = false;
                GlobalVars.highlightCurrent = false;
            }

        }

        private void Main_Form_Load(object sender, EventArgs e)
        {
            dataGridView1.ClearSelection();
        }

        private void customChrgToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "Custom Chg";
        }

        private void asReceivedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "As Received";
        }

        private void fullChargeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "Full Charge-6";
        }

        private void fullCharge4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "Full Charge-4";
        }

        private void topCharge4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "Top Charge-4";
        }

        private void topCharge2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "Top Charge-2";
        }

        private void topCharge1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "Top Charge-1";
        }

        private void capacity1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "Capacity-1";
        }

        private void dischargeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "Discharge";
        }

        private void slowCharge14ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "Slow Charge-14";
        }

        private void slowCharge16ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "Slow Charge-16";
        }

        private void testToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "Test";
        }

        private void customCapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "Custom Cap";
        }

        private void reflexChg1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "Reflex Chg-1";
        }

        private void clearToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "";
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "";
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "0";
        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "1";
        }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "2";
        }

        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "3";
        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "4";
        }

        private void toolStripMenuItem12_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "5";
        }

        private void toolStripMenuItem13_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "6";
        }

        private void toolStripMenuItem14_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "7";
        }

        private void toolStripMenuItem15_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "8";
        }

        private void toolStripMenuItem16_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "9";
        }

        private void toolStripMenuItem17_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "10";
        }

        private void toolStripMenuItem18_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "11";
        }

        private void toolStripMenuItem19_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "12";
        }

        private void toolStripMenuItem20_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "13";
        }

        private void toolStripMenuItem21_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "14";
        }

        private void toolStripMenuItem22_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][9] = "15";
        }

        private void masterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Master Selected.  Needs to be implemented...");
        }

        private void slaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Slave Selected.  Needs to be implemented...");
        }

        private void cMSChargerType_Opening(object sender, CancelEventArgs e)
        {

        }

        private void cCAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][10] = "CCA";
        }

        private void iCAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][10] = "ICA";
        }

        private void otherToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][10] = "Other";
        }

        private void clearToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][10] = "";
        }

        private void cMSStartStop_Opening(object sender, CancelEventArgs e)
        {

        }

        private void startNewTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Start New Test Selected.  Needs to be implemented...");
        }

        private void resumeTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Resume Test Selected.  Needs to be implemented...");
        }

        private void stopTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Stop Test Selected.  Needs to be implemented...");
        }

        private void viewEditDeleteWorkOrdersToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void databindingTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void viewEditDeleteCustomersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmVECustomers f2 = new frmVECustomers();
            f2.Show();
        }

        private void customersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is frmVECustomers)
                {
                    return;
                }
            }
            frmVECustomers f2 = new frmVECustomers();
            f2.Show();
        }

        private void viewStandardBatteriesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is frmVStandardBats)
                {
                    return;
                }
            }
            frmVStandardBats f2 = new frmVStandardBats();
            f2.Show();

        }

        private void editTechniciansToolStripMenuItem_Click(object sender, EventArgs e)
        {

            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is frmVETechs)
                {
                    return;
                }
            }
            frmVETechs f2 = new frmVETechs();
            f2.Owner = this;
            f2.Show();
           
        }

        private void viewEditDeleteBatteriesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is frmVECustomBats)
                {
                    return;
                }
            }
            frmVECustomBats f2 = new frmVECustomBats();
            f2.Show();
        }

        private void customerBatteriesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is frmVECustomerBats)
                {
                    return;
                }
            }
            frmVECustomerBats f2 = new frmVECustomerBats();
            f2.Show();

        }

        private void batteriesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is frmVECustomBats)
                {
                    return;
                }
            }
            frmVECustomBats f2 = new frmVECustomBats();
            f2.Show();
        }

        private void workOrdersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is frmVEWorkOrders)
                {
                    return;
                }
            }
            frmVEWorkOrders f2 = new frmVEWorkOrders();
            f2.Show();
        }




    }
}
