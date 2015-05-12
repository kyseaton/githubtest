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

                InitializeGrid();
                InitializeTimers();
                Scan();

            }
            catch(Exception ex)
            {
                MessageBox.Show("In Main_Form:  " + ex.ToString());
            }


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
            Business_Name bn = new Business_Name(this);
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
                    });

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
                for (int i = 0; i < 16; i++)
                {

                    this.Invoke((MethodInvoker)delegate
                    {
                        dataGridView1.CurrentCell = dataGridView1.Rows[i].Cells[0];
                        dataGridView1.ClearSelection();
                    });

                    Thread.Sleep(1000);
                    this.Invoke((MethodInvoker)delegate
                    {
                        if (dataGridView1.Rows[i].Cells[4].Style.BackColor == Color.Red)
                        {
                            d.Rows[i][4] = false;
                            dataGridView1.Rows[i].Cells[4].Style.BackColor = Color.Gainsboro;
                        }
                    });
                }

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



    }
}
