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
using System.IO;


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

        /// <summary>
        /// This function looks at the DB and fills up the dropdown designating the oporator
        /// </summary>
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
            comboBox1.SelectedValue = GlobalVars.currentTech;


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


            if (GlobalVars.autoConfig) { this.automaticallyConfigureChargerToolStripMenuItem.Checked = true; }
            else { this.automaticallyConfigureChargerToolStripMenuItem.Checked = false; }

            this.comboBox1.SelectedValue = GlobalVars.currentTech;

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
                MessageBox.Show("Error: Failed to store new data in the DataBase.\n" + ex.Message);
                return;
            }

            // We also have to save our Com Ports!
            try
            {
                strUpdateCMD = "UPDATE Comconfig SET Comm1='" + GlobalVars.CSCANComPort + "', Comm2='" + GlobalVars.ICComPort + "';";
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                myAccessConn.Open();
                myAccessCommand.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to store new data in the DataBase.\n" + ex.Message);
                return;
            }

            // And the auto config settings!
            try
            {
                strUpdateCMD = "UPDATE ProgramSettings SET SettingValue='" + GlobalVars.autoConfig.ToString() + "' WHERE SettingName='AutoConfigCharger';";
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                myAccessConn.Open();
                myAccessCommand.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to store new data in the DataBase.\n" + ex.Message);
                return;
            }

            // And the current tech settings!
            try
            {
                strUpdateCMD = "UPDATE ProgramSettings SET SettingValue='" + GlobalVars.currentTech + "' WHERE SettingName='CurrentTech';";
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                myAccessConn.Open();
                myAccessCommand.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to store new data in the DataBase.\n" + ex.Message);
                return;
            }

            finally
            {
                myAccessConn.Close();
            }

        }

        private void bussinessNameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is Business_Name)
                {
                    return;
                }
            }

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
            updateD(channel,1,workOrder);
            // clear the grid
            updateD(channel, 2, "");
            updateD(channel, 3, "");
            updateD(channel, 6, "");
            updateD(channel, 7, "");
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
            //save the grid for the next time we restart
            using (StreamWriter writer = new StreamWriter("../main_grid.xml",false))
            {
                d.WriteXml(writer);
            }
            

            // tell those threadpool work items to stop!!!!!
            try
            {
                cPollIC.Cancel();
                cPollCScans.Cancel();
                sequentialScanT.Cancel();
                // make sure it takes...
                Thread.Sleep(500);  
            }
            catch(Exception ex)
            {
                if (ex is NullReferenceException)
                {

                }
                else
                {
                    throw ex;
                }

            }
    
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
            updateD(dataGridView1.CurrentRow.Index,2,"Custom Chg");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void asReceivedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "As Received");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void fullChargeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Full Charge-6");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void fullCharge4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Full Charge-4");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void topCharge4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Top Charge-4");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void topCharge2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Top Charge-2");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void topCharge1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Top Charge-1");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void capacity1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Capacity-1");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void dischargeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Discharge");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void slowCharge14ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Slow Charge-14");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void slowCharge16ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Slow Charge-16");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void testToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Test");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void customCapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Custom Cap");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void constantVoltageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Constant Voltage");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void clearToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            d.Rows[dataGridView1.CurrentRow.Index][2] = "";
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string current = d.Rows[dataGridView1.CurrentRow.Index][9].ToString();

            if (current.Length > 2)
            {
                //we've got a master or a slave
                // check for slaves or master associated with this channel also

                if (current.Length == 3)
                {
                    // one digit case
                    current = current.Substring(0, 1);
                }
                else
                {
                    // two digit case
                    current = current.Substring(0, 2);
                }


                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString() == "")
                    {
                        //go to the next
                        ;
                    }
                    else if (d.Rows[i][9].ToString().Substring(0,current.Length) == current)
                    {
                        // found it!
                        // make that one the master
                        d.Rows[i][9] = current;
                        // Now enable adding another...
                        switch (Convert.ToInt32(current))
                        {
                            case 0:
                                toolStripMenuItem7.Enabled = true;
                                break;
                            case 1:
                                toolStripMenuItem8.Enabled = true;
                                break;
                            case 2:
                                toolStripMenuItem9.Enabled = true;
                                break;
                            case 3:
                                toolStripMenuItem10.Enabled = true;
                                break;
                            case 4:
                                toolStripMenuItem11.Enabled = true;
                                break;
                            case 5:
                                toolStripMenuItem12.Enabled = true;
                                break;
                            case 6:
                                toolStripMenuItem13.Enabled = true;
                                break;
                            case 7:
                                toolStripMenuItem14.Enabled = true;
                                break;
                            case 8:
                                toolStripMenuItem15.Enabled = true;
                                break;
                            case 9:
                                toolStripMenuItem16.Enabled = true;
                                break;
                            case 10:
                                toolStripMenuItem17.Enabled = true;
                                break;
                            case 11:
                                toolStripMenuItem18.Enabled = true;
                                break;
                            case 12:
                                toolStripMenuItem19.Enabled = true;
                                break;
                            case 13:
                                toolStripMenuItem20.Enabled = true;
                                break;
                            case 14:
                                toolStripMenuItem21.Enabled = true;
                                break;
                            case 15:
                                toolStripMenuItem22.Enabled = true;
                                break;
                        }// end switch

                    }// end else if
                }

            }
            
            // we always clear the current one..
            d.Rows[dataGridView1.CurrentRow.Index][9] = "";

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");

        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if ((string) d.Rows[i][9] == "0" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "0-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "0-S";
                    // Now disable adding another...
                    toolStripMenuItem7.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "0";

            // also check for a charger if the channel is linked...
            if ((bool) d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");

        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "1" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "1-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "1-S";
                    // Now disable adding another...
                    toolStripMenuItem8.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "1";
            
            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "2" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "2-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "2-S";
                    // Now disable adding another...
                    toolStripMenuItem9.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "2";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "3" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "3-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "3-S";
                    // Now disable adding another...
                    toolStripMenuItem10.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "3";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "4" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "4-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "4-S";
                    // Now disable adding another...
                    toolStripMenuItem11.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "4";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem12_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "5" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "5-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "5-S";
                    // Now disable adding another...
                    toolStripMenuItem12.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "5";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem13_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "6" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "6-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "6-S";
                    // Now disable adding another...
                    toolStripMenuItem13.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "6";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem14_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "7" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "7-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "7-S";
                    // Now disable adding another...
                    toolStripMenuItem14.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "7";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem15_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "8" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "8-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "8-S";
                    // Now disable adding another...
                    toolStripMenuItem15.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "8";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem16_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "9" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "9-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "9-S";
                    // Now disable adding another...
                    toolStripMenuItem16.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "9";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem17_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "10" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "10-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "10-S";
                    // Now disable adding another...
                    toolStripMenuItem17.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "10";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem18_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "11" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "11-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "11-S";
                    // Now disable adding another...
                    toolStripMenuItem18.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "11";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem19_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "12" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "12-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "12-S";
                    // Now disable adding another...
                    toolStripMenuItem19.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "12";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem20_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "13" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "13-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "13-S";
                    // Now disable adding another...
                    toolStripMenuItem20.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "13";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem21_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "14" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "14-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "14-S";
                    // Now disable adding another...
                    toolStripMenuItem21.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "14";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void toolStripMenuItem22_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "15" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    d.Rows[i][9] = "15-M";
                    // and the current one the slave
                    d.Rows[dataGridView1.CurrentRow.Index][9] = "15-S";
                    // Now disable adding another...
                    toolStripMenuItem22.Enabled = false;
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            d.Rows[dataGridView1.CurrentRow.Index][9] = "15";

            // also check for a charger if the channel is linked...
            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][8] == true)
            {
                checkForIC(int.Parse((string)d.Rows[dataGridView1.CurrentRow.Index][9]), dataGridView1.CurrentRow.Index);
            }

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
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
            updateD(dataGridView1.CurrentRow.Index,10,"CCA");

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void iCAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 10, "ICA");

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void otherToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 10, "Other");

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void clearToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 10, "");

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
        }

        private void cMSStartStop_Opening(object sender, CancelEventArgs e)
        {

        }

        private void startNewTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //clear the E-time to set the test to a new test
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            // we will run the tests on a helper thread
            // helper thread code is located in Test_Portion.cs
            RunTest();

        }

        private void resumeTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //runtest without resetting the time!
            RunTest();
        }

        private void stopTestToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cRunTest[dataGridView1.CurrentRow.Index].Cancel(); 
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

        private void commPortSettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is ComportSettings)
                {
                    return;
                }
            }
            ComportSettings f2 = new ComportSettings();
            f2.Owner = this;
            f2.Show();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is ICSettingsForm)
                {
                    return;
                }
            }
            ICSettingsForm f2 = new ICSettingsForm();
            f2.Owner = this;
            f2.Show();

        }

        private void automaticallyConfigureChargerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (automaticallyConfigureChargerToolStripMenuItem.Checked == false)
            {
                automaticallyConfigureChargerToolStripMenuItem.Checked = true;
                GlobalVars.autoConfig = true;
            }
            else
            {
                automaticallyConfigureChargerToolStripMenuItem.Checked = false;
                GlobalVars.autoConfig = false;
            }
        }

        private void chargerConfigurationInterfaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is ICSettingsForm)
                {
                    return;
                }
            }
            ICSettingsForm f2 = new ICSettingsForm();
            f2.Owner = this;
            f2.Show();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            GlobalVars.currentTech = (string)comboBox1.SelectedValue;
        }

        private void cMSChargerChannel_Opening(object sender, CancelEventArgs e)
        {

        }



  




    }
}
