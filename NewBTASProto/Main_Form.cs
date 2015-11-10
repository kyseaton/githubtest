﻿using System;
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
using System.Net.Mail;



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
                Initialize_Graph_Settings();
                Initialize_PCI_Settings();

                InitializeGrid();
                InitializeTimers();
                Scan();

                SetChargersCriticalAtStart();

                GlobalVars.loading = false;

                this.Height = (int) Properties.Settings.Default.FormHeight;
                this.Width = (int) Properties.Settings.Default.FormWidth;

       
            }
            catch(Exception ex)
            {
                MessageBox.Show("In Main_Form:  " + ex.ToString());
            }


        }

        private void SetChargersCriticalAtStart()
        {
            // loop through the grid and set chargers to critical (look for ICs) if they are assigned
            for (int i = 0; i < 16; i++)
            {
                //if the charger is linked and there is a number assigned and not a slave make the charger critical
                if ((bool)d.Rows[i][8] == true && d.Rows[i][9].ToString() != "" && !(d.Rows[i][9].ToString().Contains("s")))
                {
                    if (d.Rows[i][9].ToString().Length < 3)
                    {
                        criticalNum[int.Parse(d.Rows[i][9].ToString())] = true;
                    }
                    else if (d.Rows[i][9].ToString().Length == 3)
                    {
                        criticalNum[int.Parse(d.Rows[i][9].ToString().Substring(0,1))] = true;
                    }
                    else
                    {
                        criticalNum[int.Parse(d.Rows[i][9].ToString().Substring(0, 2))] = true;
                    }
                }
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
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
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

                lock (dataBaseLock)
                {
                    myAccessConn.Open();
                    myDataAdapter.Fill(operators, "Operators");
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

            toolStripStatusLabel4.Text = "Version:  " + GlobalVars.programVersion;

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
            Graphics_Form gf = new Graphics_Form(d.Rows[dataGridView1.CurrentRow.Index][3].ToString(), d.Rows[dataGridView1.CurrentRow.Index][1].ToString());
            gf.Owner = this;
            gf.Show();
   
        }

        private void testToolStripMenuItem_Click(object sender, EventArgs e)
        {


            Reports_Form rf = new Reports_Form(d.Rows[dataGridView1.CurrentRow.Index][3].ToString(), d.Rows[dataGridView1.CurrentRow.Index][1].ToString());
            rf.Owner = this;
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
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
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

                lock (dataBaseLock)
                {
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                }

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

                lock (dataBaseLock)
                {
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                }
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

                lock (dataBaseLock)
                {
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                }
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

                lock (dataBaseLock)
                {
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to store new data in the DataBase.\n" + ex.Message);
                return;
            }

            finally
            {
                lock (dataBaseLock)
                {
                    myAccessConn.Close();
                }
            }

        }

        private void bussinessNameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is Business_Name)
                {
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
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
            // clear the grid (if it's not on a slave channel...
            if (!d.Rows[channel][9].ToString().Contains("S"))
            {
                updateD(channel, 2, "");
                updateD(channel, 3, "");
                updateD(channel, 6, "");
                updateD(channel, 7, "");
            }
            if (d.Rows[channel][9].ToString().Contains("M"))
            {
                // find the slave and update it also!
                //find the slave
                string temp = d.Rows[channel][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }

            }

            // also re set the combos...
            fillPlotCombos(channel);

            // finally we need to update the pci datatable
            if (workOrder != "")
            {
                //split off the first work order if we have multiple ones..
                string tempWOS = workOrder;
                char[] delims = { ' ' };
                string[] A = tempWOS.Split(delims);
                workOrder = A[0];

                DataSet batData = new DataSet();

                // find out the nominal voltage 
                // first get the battery Model from the work order..
                // Open database containing all the battery data....
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                string strAccessSelect = @"SELECT BatteryModel FROM WorkOrders WHERE WorkOrderNumber='" + workOrder.Trim() + @"'";

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

                    lock (dataBaseLock)
                    {
                        myAccessConn.Open();
                        myDataAdapter.Fill(batData, "Bat");
                        myAccessConn.Close();
                    }


                }
                catch (Exception ex)
                {
                    myAccessConn.Close();
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    return;
                }

                //Now we have the battery Model...
                pci.Rows[channel][0] = batData.Tables[0].Rows[0][0].ToString();

                // Lets get the nominal voltage!
                strAccessSelect = @"SELECT BTECH,VOLT,NCELLS,BCVMIN,BCVMAX,CCVMMIN,CCVMAX,CCAPV FROM BatteriesCustom WHERE BatteryModel='" + batData.Tables[0].Rows[0][0].ToString() + @"'";
                batData = new DataSet();
                //  now try to access it

                try
                {
                    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                    lock (dataBaseLock)
                    {
                        myAccessConn.Open();
                        myDataAdapter.Fill(batData, "Bat");
                        myAccessConn.Close();
                    }

                }
                catch (Exception ex)
                {
                    myAccessConn.Close();
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    return;
                }

                // we should have the data!
                // now put it into the the pci!
                if (batData.Tables[0].Rows[0][0].ToString() != "")
                {
                    pci.Rows[channel][1] = batData.Tables[0].Rows[0][0].ToString();
                }
                if (batData.Tables[0].Rows[0][1].ToString() != "")
                {
                    pci.Rows[channel][2] = float.Parse(batData.Tables[0].Rows[0][1].ToString());
                }
                if (batData.Tables[0].Rows[0][2].ToString() != "")
                {
                    pci.Rows[channel][3] = int.Parse(batData.Tables[0].Rows[0][2].ToString()); 
                }
                if (batData.Tables[0].Rows[0][3].ToString() != "")
                {
                    pci.Rows[channel][4] = float.Parse(batData.Tables[0].Rows[0][3].ToString());
                }
                if (batData.Tables[0].Rows[0][4].ToString() != "")
                {
                    pci.Rows[channel][5] = float.Parse(batData.Tables[0].Rows[0][4].ToString());
                }
                if (batData.Tables[0].Rows[0][5].ToString() != "")
                {
                    pci.Rows[channel][6] = float.Parse(batData.Tables[0].Rows[0][5].ToString());
                }
                else if (batData.Tables[0].Rows[0][5].ToString() == "" && batData.Tables[0].Rows[0][0].ToString() == "NiCd ULM")
                {
                    pci.Rows[channel][6] = 1.82;
                }
                if (batData.Tables[0].Rows[0][6].ToString() != "")
                {
                    pci.Rows[channel][7] = float.Parse(batData.Tables[0].Rows[0][6].ToString());
                }
                if (batData.Tables[0].Rows[0][7].ToString() != "")
                {
                    pci.Rows[channel][8] = float.Parse(batData.Tables[0].Rows[0][7].ToString());
                }
            }
            else
            {
                // we don't have a workorder
                // reset to default...
                pci.Rows[channel][0] = "None";
                pci.Rows[channel][1] = "NiCd";
                pci.Rows[channel][2] = 24;         // negative 1 is the default...
                pci.Rows[channel][3] = -1;         // negative 1 is the default...
                pci.Rows[channel][4] = -1;         // negative 1 is the default...
                pci.Rows[channel][5] = -1;         // negative 1 is the default...
                pci.Rows[channel][6] = -1;         // negative 1 is the default...
                pci.Rows[channel][7] = 1.75;         // negative 1 is the default...
                pci.Rows[channel][8] = -1;         // negative 1 is the default...
            }
        }


        private void btnGetSerialPorts_Click_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }


        public CancellationTokenSource cFindStations = new CancellationTokenSource();

        private void button1_Click(object sender, EventArgs e)
        {

            // create a thread to go through and look for the stations, this way the UI will still work while the search is happening
            ThreadPool.QueueUserWorkItem(s =>
            {

                // setup the canellation token
                CancellationToken token = (CancellationToken)s;


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

                Thread.Sleep(250);
                if (token.IsCancellationRequested) { return; }
                Thread.Sleep(250);
                if (token.IsCancellationRequested) { return; }

                // turn on all of the in use buttons
                for (int i = 0; i < 16; i++)
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        updateD(i,4, true);
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
                    Thread.Sleep(250);
                    if (token.IsCancellationRequested) { return; }
                    Thread.Sleep(250);
                    if (token.IsCancellationRequested) { return; }
                    Thread.Sleep(250);
                    if (token.IsCancellationRequested) { return; }
                    Thread.Sleep(250);
                    if (token.IsCancellationRequested) { return; }

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
                            updateD(i,4, false);
                            dataGridView1.Rows[i].Cells[4].Style.BackColor = Color.Gainsboro;
                        }
                    });
                }

                //Finally take care of the last channel
                //give it time to check the channel

                Thread.Sleep(250);
                if (token.IsCancellationRequested) { return; }
                Thread.Sleep(250);
                if (token.IsCancellationRequested) { return; }
                Thread.Sleep(250);
                if (token.IsCancellationRequested) { return; }
                Thread.Sleep(250);
                if (token.IsCancellationRequested) { return; }

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
                        updateD(15,4, false);
                        dataGridView1.Rows[15].Cells[4].Style.BackColor = Color.Gainsboro;
                    }
                });

                bool foundOne = false;
                // let's see if we found any!
                for (int i = 0; i < 16; i++)
                {
                    if ((bool)d.Rows[i][4])
                    {
                        foundOne = true;
                        break;
                    }
                }

                if (!foundOne)
                {
                    // flip the comms!
                    // stop all of the scanning threads
                    try
                    {
                        cPollIC.Cancel();
                        cPollCScans.Cancel();
                        sequentialScanT.Cancel();

                        cPollIC.Dispose();
                        cPollCScans.Dispose();
                        sequentialScanT.Dispose();
                    }
                    catch (Exception ex)
                    {
                        if (ex is NullReferenceException || ex is ObjectDisposedException)
                        {

                        }
                        else
                        {
                            throw ex;
                        }
                    }


                    // close the comms
                    CSCANComPort.Close();
                    ICComPort.Close();

                    //Update the Globals
                    string temp = GlobalVars.CSCANComPort;
                    GlobalVars.CSCANComPort = GlobalVars.ICComPort;
                    GlobalVars.ICComPort = temp;

                    //Start the threads back up
                    Scan();

                    //rerun the same code...
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

                    Thread.Sleep(250);
                    if (token.IsCancellationRequested) { return; }
                    Thread.Sleep(250);
                    if (token.IsCancellationRequested) { return; }


                    // turn on all of the in use buttons
                    for (int i = 0; i < 16; i++)
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            updateD(i, 4, true);
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
                        Thread.Sleep(250);
                        if (token.IsCancellationRequested) { return; }
                        Thread.Sleep(250);
                        if (token.IsCancellationRequested) { return; }
                        Thread.Sleep(250);
                        if (token.IsCancellationRequested) { return; }
                        Thread.Sleep(250);
                        if (token.IsCancellationRequested) { return; }

                        // move the current channel
                        this.Invoke((MethodInvoker)delegate
                        {
                            dataGridView1.CurrentCell = dataGridView1.Rows[i + 1].Cells[0];
                            dataGridView1.ClearSelection();
                        });

                        // wait again
                        Thread.Sleep(100);
                        this.Invoke((MethodInvoker)delegate
                        {
                            if (dataGridView1.Rows[i].Cells[4].Style.BackColor == Color.Red)
                            {
                                updateD(i, 4, false);
                                dataGridView1.Rows[i].Cells[4].Style.BackColor = Color.Gainsboro;
                            }
                        });
                    }

                    //Finally take care of the last channel
                    //give it time to check the channel
                    Thread.Sleep(250);
                    if (token.IsCancellationRequested) { return; }
                    Thread.Sleep(250);
                    if (token.IsCancellationRequested) { return; }
                    Thread.Sleep(250);
                    if (token.IsCancellationRequested) { return; }
                    Thread.Sleep(250);
                    if (token.IsCancellationRequested) { return; }

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
                            updateD(15, 4, false);
                            dataGridView1.Rows[15].Cells[4].Style.BackColor = Color.Gainsboro;
                        }
                    });
                }// end if

                    //reenable the button before exit
                this.Invoke((MethodInvoker)delegate
                {
                    // start by disabling the button while we look for stations
                    button1.Enabled = true;
                    // also disable the grid, so the user cannot interfere with the search
                    dataGridView1.Enabled = true;
                });

            },cFindStations.Token);                     // end thread



        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void Main_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            // We need to check if there are tests running and ask the user if they are sure they want to quite in the event that there are tests running...
            for (int i = 0; i < 16; i++)
            {
                if ((bool) d.Rows[i][5] == true)
                {
                    DialogResult dialogResult = MessageBox.Show("There is a test running. If you quit, the test data will no longer be recorded. You will also need to attend to the charger associated with the test, as it will no longer be computer controlled.", "Are you sure you want to quit?", MessageBoxButtons.OKCancel);
                    if (dialogResult == DialogResult.OK)
                    {
                        break;
                    }
                    else 
                    {
                        e.Cancel = true;
                        return; 
                    }

                }
            }// end for


            //save the grid for the next time we restart
            using (StreamWriter writer = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\main_grid.xml",false))
            {
                for (int i = 0; i < 16; i++)
                {
                    updateD(i, 5, false);
                    updateD(i, 6, "");
                    updateD(i, 7, "");
                    updateD(i, 10, "");
                    updateD(i, 11, "");
                }// end for

                d.WriteXml(writer);
            }

            //save the grid for the next time we restart
            using (StreamWriter writer = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\graph_set.xml", false))
            {
                gs.WriteXml(writer);
            }

            //save the pci grid for the next time we restart
            using (StreamWriter writer = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\pci_set.xml", false))
            {
                pci.WriteXml(writer);
            }
            
            //Save the current form width and height
            if (this.WindowState == FormWindowState.Maximized)
            {
                Properties.Settings.Default.maximized = true;
            }
            else
            {
                Properties.Settings.Default.maximized = false;
                Properties.Settings.Default.FormHeight = this.Height;
                Properties.Settings.Default.FormWidth = this.Width;
                Properties.Settings.Default.PositionX = this.Location.X;
                Properties.Settings.Default.PositionY = this.Location.Y;
            }

            Properties.Settings.Default.Save();
            
            // tell those threadpool work items to stop!!!!!
            try
            {
                cPollIC.Cancel();
                cPollCScans.Cancel();
                sequentialScanT.Cancel();
                cFindStations.Cancel();
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
            if ((bool)Properties.Settings.Default.maximized == true)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else if(Properties.Settings.Default.PositionX > 100 && Properties.Settings.Default.PositionY > 100)
            {
                this.Location = new Point((int) Properties.Settings.Default.PositionX, (int)Properties.Settings.Default.PositionY);
            }
            
        }

        private void customChrgToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index,2,"Custom Chg");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "Custom Chg");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }
        }

        private void asReceivedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "As Received");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "As Received");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }

        }

        private void fullChargeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Full Charge-6");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "Full Charge-6");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }

        }

        private void fullCharge4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Full Charge-4");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "Full Charge-4");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }
        }

        private void topCharge4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Top Charge-4");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "Top Charge-4");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }
        }

        private void topCharge2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Top Charge-2");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "Top Charge-2");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }
        }

        private void topCharge1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Top Charge-1");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "Top Charge-1");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }
        }

        private void capacity1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Capacity-1");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "Capacity-1");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }
        }

        private void dischargeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Discharge");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "Discharge");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }
        }

        private void slowCharge14ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Slow Charge-14");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "Slow Charge-14");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }
        }

        private void slowCharge16ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Slow Charge-16");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "Slow Charge-16");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }
        }

        private void testToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Test");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "Test");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }
        }

        private void customCapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Custom Cap");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "Custom Cap");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }
        }

        private void constantVoltageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "Constant Voltage");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "Constant Voltage");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }
        }

        private void clearToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            updateD(dataGridView1.CurrentRow.Index, 2, "");
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);

            // also update the slave (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        updateD(i, 2, "");
                        updateD(i, 3, "");
                        updateD(i, 6, "");
                        updateD(i, 7, "");
                    }
                }
            }
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // first update the slave colors (if we have a master...)
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                //find the slave
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        // also change the grid color
                        dataGridView1.Rows[i].Cells[2].Style.BackColor = Color.Aquamarine;
                        dataGridView1.Rows[i].Cells[5].Style.BackColor = Color.Gainsboro;
                        dataGridView1.Rows[i].Cells[8].Style.BackColor = Color.Gainsboro;
                        dataGridView1.Rows[i].Cells[12].Style.BackColor = Color.LightSkyBlue;
                    }
                }
            }
            else if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("S"))
            {
                // also change the grid color
                dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.Aquamarine;
                dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.Gainsboro;
                dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.Gainsboro;
                dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSkyBlue;
                updateD(dataGridView1.CurrentRow.Index, 12, false);
            }

            //Now onto the normal stuff...
            correctMasterSlave();
            // we always clear the current one..
            updateD(dataGridView1.CurrentRow.Index,9, "");

            //make sure we clear the current test
            updateD(dataGridView1.CurrentRow.Index, 3, "");
            updateD(dataGridView1.CurrentRow.Index, 6, "");
            updateD(dataGridView1.CurrentRow.Index, 7, "");
            fillPlotCombos(dataGridView1.CurrentRow.Index);


        }

        private void correctMasterSlave()
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
                    if (d.Rows[i][9].ToString() == "" || d.Rows[i][9].ToString().Length != (current.Length + 2))
                    {
                        //go to the next
                        ;
                    }
                    else if (d.Rows[i][9].ToString().Substring(0, current.Length) == current)
                    {
                        // found it!
                        // make that one the master
                        updateD(i,9, current);
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
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if ((string) d.Rows[i][9] == "0" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9,"0-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9,"0-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem7.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9,"0");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "1" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "1-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "1-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem8.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "1");
            
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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "2" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9,"2-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "2-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem9.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "2");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "3" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "3-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "3-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem10.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "3");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "4" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "4-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "4-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem11.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "4");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "5" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "5-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "5-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem12.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "5");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "6" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "6-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "6-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem13.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "6");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "7" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "7-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "7-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem14.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "7");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "8" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "8-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "8-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem15.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "8");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "9" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "9-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "9-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem16.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "9");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "10" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "10-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "10-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem17.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "10");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "11" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "11-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "11-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem18.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "11");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "12" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "12-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "12-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem19.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "12");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "13" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "13-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "13-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem20.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "13");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "14" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "14-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "14-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;
                    // Now disable adding another...
                    toolStripMenuItem21.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "14");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString() != "") { correctMasterSlave(); }

            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString() == "15" && i != dataGridView1.CurrentRow.Index)
                {
                    // there is already a zero in one of the other rows!
                    // make that one the master
                    updateD(i,9, "15-M");
                    // and the current one the slave
                    updateD(dataGridView1.CurrentRow.Index,9, "15-S");
                    // also change the grid color
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[12].Style.BackColor = Color.LightSteelBlue;

                    // Now disable adding another...
                    toolStripMenuItem22.Enabled = false;
                    // also syncronyze them
                    d.Rows[dataGridView1.CurrentRow.Index][8] = d.Rows[i][8];
                    d.Rows[dataGridView1.CurrentRow.Index][12] = d.Rows[i][12];
                    d.Rows[dataGridView1.CurrentRow.Index][2] = d.Rows[i][2];
                    d.Rows[dataGridView1.CurrentRow.Index][10] = d.Rows[i][10];
                    d.Rows[dataGridView1.CurrentRow.Index][11] = d.Rows[i][11];
                    if (dataGridView1.Rows[i].Cells[8].Style.BackColor != Color.Gainsboro)
                    {
                        dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = dataGridView1.Rows[i].Cells[8].Style.BackColor;
                    }
                    return;
                }
            }
            // otherwise we'll proceed as normal...
            updateD(dataGridView1.CurrentRow.Index,9, "15");

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
            updateD(dataGridView1.CurrentRow.Index, 10, "Shunt");

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
            if (d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Contains("M"))
            {
                // we need to find the slave and clear it also...
                string temp = d.Rows[dataGridView1.CurrentRow.Index][9].ToString().Replace("-M", "");

                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        //found the slave
                        updateD(i, 6, "");
                        break;
                    }
                }
            }
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
            try
            {
                cRunTest[dataGridView1.CurrentRow.Index].Cancel();
            }
            catch
            {
                updateD(dataGridView1.CurrentRow.Index, 5, false);
            }
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

        private void customersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is frmVECustomers)
                {
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
                    frm.BringToFront();
                    return;
                }
            }
            frmVECustomers f2 = new frmVECustomers();
            f2.Owner = this;
            f2.Show();
        }


        private void editTechniciansToolStripMenuItem_Click(object sender, EventArgs e)
        {

            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is frmVETechs)
                {
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
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
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
                    return;
                }
            }
            frmVECustomBats f2 = new frmVECustomBats();
            f2.Owner = this;
            f2.Show();
        }

        private void customerBatteriesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is frmVECustomerBats)
                {
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
                    return;
                }
            }
            frmVECustomerBats f2 = new frmVECustomerBats();
            f2.Owner = this;
            f2.Show();

        }

        private void batteriesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is frmVECustomBats)
                {
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
                    return;
                }
            }
            frmVECustomBats f2 = new frmVECustomBats();
            f2.Owner = this;
            f2.Show();
        }

        private void workOrdersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is frmVEWorkOrders)
                {
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
                    return;
                }
            }
            frmVEWorkOrders f2 = new frmVEWorkOrders();
            f2.Owner = this;
            f2.Show();
        }

        private void commPortSettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is ComportSettings)
                {
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
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
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
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

            dataGridView1_Resize(this, null);
        }

        private void chargerConfigurationInterfaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is ICSettingsForm)
                {
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
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

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {

        }

        private void rtbIncoming_Resize(object sender, EventArgs e)
        {


        }

        private void Main_Form_ResizeEnd(object sender, EventArgs e)
        {

        }

        private void dataGridView1_Resize(object sender, EventArgs e)
        {
            int cumWidth = 0;
            //Scale the columns to the new width!
            if (GlobalVars.autoConfig)
            {
                dataGridView1.Columns[0].Width = (40 * dataGridView1.Width) / 1057;
                cumWidth += (40 * dataGridView1.Width) / 1057;
                dataGridView1.Columns[1].Width = (180 * dataGridView1.Width) / 1057;
                cumWidth += (180 * dataGridView1.Width) / 1057;
                dataGridView1.Columns[2].Width = (140 * dataGridView1.Width) / 1057;
                cumWidth += (140 * dataGridView1.Width) / 1057;
                dataGridView1.Columns[3].Width = (40 * dataGridView1.Width) / 1057;
                cumWidth += (40 * dataGridView1.Width) / 1057;
                dataGridView1.Columns[4].Width = (44 * dataGridView1.Width) / 1057;
                cumWidth += (44 * dataGridView1.Width) / 1057;
                dataGridView1.Columns[5].Width = (44 * dataGridView1.Width) / 1057;
                cumWidth += (44 * dataGridView1.Width) / 1057;
                dataGridView1.Columns[6].Width = (100 * dataGridView1.Width) / 1057;
                cumWidth += (100 * dataGridView1.Width) / 1057;
                dataGridView1.Columns[7].Width = (120 * dataGridView1.Width) / 1057;
                cumWidth += (120 * dataGridView1.Width) / 1057;
                dataGridView1.Columns[8].Width = (60 * dataGridView1.Width) / 1057;
                cumWidth += (60 * dataGridView1.Width) / 1057;
                dataGridView1.Columns[9].Width = (50 * dataGridView1.Width) / 1057;
                cumWidth += (50 * dataGridView1.Width) / 1057;
                dataGridView1.Columns[10].Width = (78 * dataGridView1.Width) / 1057;
                cumWidth += (78 * dataGridView1.Width) / 1057;
                dataGridView1.Columns[11].Width = (78 * dataGridView1.Width) / 1057;
                cumWidth += (78 * dataGridView1.Width) / 1057;
                dataGridView1.Columns[12].Width = (dataGridView1.Width - 43) - cumWidth;
            }
            else
            {
                dataGridView1.Columns[0].Width = (40 * dataGridView1.Width) / 1017;
                cumWidth += (40 * dataGridView1.Width) / 1017;
                dataGridView1.Columns[1].Width = (180 * dataGridView1.Width) / 1017;
                cumWidth += (180 * dataGridView1.Width) / 1017;
                dataGridView1.Columns[2].Width = (140 * dataGridView1.Width) / 1017;
                cumWidth += (140 * dataGridView1.Width) / 1017;
                dataGridView1.Columns[3].Width = (40 * dataGridView1.Width) / 1017;
                cumWidth += (40 * dataGridView1.Width) / 1017;
                dataGridView1.Columns[4].Width = (44 * dataGridView1.Width) / 1017;
                cumWidth += (44 * dataGridView1.Width) / 1017;
                dataGridView1.Columns[5].Width = (44 * dataGridView1.Width) / 1017;
                cumWidth += (44 * dataGridView1.Width) / 1017;
                dataGridView1.Columns[6].Width = (100 * dataGridView1.Width) / 1017;
                cumWidth += (100 * dataGridView1.Width) / 1017;
                dataGridView1.Columns[7].Width = (120 * dataGridView1.Width) / 1017;
                cumWidth += (120 * dataGridView1.Width) / 1017;
                dataGridView1.Columns[8].Width = (60 * dataGridView1.Width) / 1017;
                cumWidth += (60 * dataGridView1.Width) / 1017;
                dataGridView1.Columns[9].Width = (50 * dataGridView1.Width) / 1017;
                cumWidth += (50 * dataGridView1.Width) / 1017;
                dataGridView1.Columns[10].Width = (78 * dataGridView1.Width) / 1017;
                cumWidth += (78 * dataGridView1.Width) / 1017;
                dataGridView1.Columns[11].Width = (dataGridView1.Width - 43) - cumWidth;
            }

        }

        private void editTestSettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is frmVETests)
                {
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
                    return;
                }
            }
            frmVETests f2 = new frmVETests();
            f2.Owner = this;
            f2.Show();
        }

        private void programVersionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is Program_Version)
                {
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
                    return;
                }
            }

            Program_Version bn = new Program_Version();
            bn.Owner = this;
            bn.Show();
        }

        private void helpToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is Help)
                {
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
                    return;
                }
            }

            Help bn = new Help();
            bn.Owner = this;
            bn.Show();

        }

        private void notificationServiceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is NoteServe)
                {
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
                    return;
                }
            }

            NoteServe bn = new NoteServe();
            bn.Owner = this;
            bn.Show();

        }


        private void sendNote(int station, int priority, string message = "Event!")
        {
            // Message Center Portion of the code
            rtbIncoming.Text = System.DateTime.Now.ToString() + ("  " + message + " (station " + station.ToString() + ")" + Environment.NewLine) + rtbIncoming.Text;

            // is the note service still on?
            if (GlobalVars.noteOn == false) { return; }

            // test if the priority is being sent
            // 1 is highpriority
            // 2 is medium
            // 3 is low

            switch (priority)
            {
                case 1:
                    break; // high priority messages always pass...
                case 2:
                    if (GlobalVars.medLev == true || GlobalVars.allLev == true) { break; } // medium make it through when medium or all is on
                    else { return; }
                case 3:
                    if (GlobalVars.allLev == true) { break; }  // low only make it through when all is on
                    else { return; }
                default:
                    return;
            }// end station switch!

            switch (station)
            {
                case 0:
                    if (GlobalVars.stat0 == false) { return; }
                    else { break; }
                case 1:
                    if (GlobalVars.stat1 == false) { return; }
                    else { break; }
                case 2:
                    if (GlobalVars.stat2 == false) { return; }
                    else { break; }
                case 3:
                    if (GlobalVars.stat3 == false) { return; }
                    else { break; }
                case 4:
                    if (GlobalVars.stat4 == false) { return; }
                    else { break; }
                case 5:
                    if (GlobalVars.stat5 == false) { return; }
                    else { break; }
                case 6:
                    if (GlobalVars.stat6 == false) { return; }
                    else { break; }
                case 7:
                    if (GlobalVars.stat7 == false) { return; }
                    else { break; }
                case 8:
                    if (GlobalVars.stat8 == false) { return; }
                    else { break; }
                case 9:
                    if (GlobalVars.stat9 == false) { return; }
                    else { break; }
                case 10:
                    if (GlobalVars.stat10 == false) { return; }
                    else { break; }
                case 11:
                    if (GlobalVars.stat11 == false) { return; }
                    else { break; }
                case 12:
                    if (GlobalVars.stat12 == false) { return; }
                    else { break; }
                case 13:
                    if (GlobalVars.stat13 == false) { return; }
                    else { break; }
                case 14:
                    if (GlobalVars.stat14 == false) { return; }
                    else { break; }
                case 15:
                    if (GlobalVars.stat15 == false) { return; }
                    else { break; }
                default:
                    return;        
            }// end station switch!


            //we made it here, so let's send a message!!!!

            // do everything on a helper thread...
            ThreadPool.QueueUserWorkItem(s =>
            {

                try
                {
                    // Create a System.Net.Mail.MailMessage object
                    MailMessage note = new MailMessage();

                    // Add a recipients
                    char[] delims = { ',' };
                    foreach (string str in GlobalVars.recipients.Split(delims))
                    {
                        if (str != "")
                        {
                            note.To.Add(str.Trim());
                        }
                    }
                    
                    // Add a message subject
                    note.Subject = "BTAS Message";

                    // Add a message body
                    note.Body = "BTAS Event (station " + station.ToString() + ") :" + message;

                    // Create a System.Net.Mail.MailAddress object and 
                    // set the sender email address and display name.
                    note.From = new MailAddress(GlobalVars.user);

                    // Create a System.Net.Mail.SmtpClient object
                    // and set the SMTP host and port number
                    SmtpClient smtp = new SmtpClient(GlobalVars.server, int.Parse(GlobalVars.port));

                    // If your server requires authentication add the below code
                    // =========================================================
                    // Enable Secure Socket Layer (SSL) for connection encryption
                    smtp.EnableSsl = true;

                    // Do not send the DefaultCredentials with requests
                    smtp.UseDefaultCredentials = false;

                    // Create a System.Net.NetworkCredential object and set
                    // the username and password required by your SMTP account
                    smtp.Credentials = new System.Net.NetworkCredential(GlobalVars.user, GlobalVars.pass);
                    // =========================================================

                    // Send the message
                    smtp.Send(note);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

            });

        }// end sendNote!

        private void backupDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string folder = "";

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                folder = folderBrowserDialog1.SelectedPath;
                // Let the user know what happned!
                try
                {
                    //try to copy the database from the appdata folder to the selected folder...

                    File.Copy(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB", folder + @"\BTAS16NV_" + System.DateTime.Now.ToString("yyyyMMddHHmmssfff") + @".MDB");
                    MessageBox.Show("Database was backed up to:  " + folder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Database was not backed up!" + Environment.NewLine + ex.ToString());
                }

            }            

        }// end backup database

        private void button5_Click_1(object sender, EventArgs e)
        {
            
        }

        private void restoreDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //first check if there is a test running and return if so
            for (int i = 0; i < 16; i++)
            {
                if ((bool)d.Rows[i][5] || d.Rows[i][2].ToString() != "")
                {
                    MessageBox.Show("Please stop all tests and clear all workorders before restoring the database!");
                    return;
                }
            }

            string folder = "";


            if (folderBrowserDialog2.ShowDialog() == DialogResult.OK)
            {

                //here we export the old DB
                folder = folderBrowserDialog2.SelectedPath;
                // Let the user know what happned!
                try
                {
                    //try to copy the database from the appdata folder to the selected folder...

                    File.Copy(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB", folder + @"\BTAS16NV_" + System.DateTime.Now.ToString("yyyyMMddHHmmssfff") + @".MDB");
                    MessageBox.Show("Database was backed up to:  " + folder);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Database was not backed up!" + Environment.NewLine + ex.ToString());
                    return;
                }

                string file;

                //here we import the new DB
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    //here we export the old DB
                    file = openFileDialog1.FileName;
                    // Let the user know what happned!
                    try
                    {
                        //try to copy the database from the appdata folder to the selected folder...

                        File.Copy(file,Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB",true);
                        MessageBox.Show("Selected database has been restored");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Database was not restored!" + Environment.NewLine + ex.ToString());
                        return;
                    }
                }

            }// end if

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void doc_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Bitmap bmp = new Bitmap(chart1.Width, chart1.Height, chart1.CreateGraphics());
            this.Invoke((MethodInvoker)delegate()
            {
                chart1.DrawToBitmap(bmp, new Rectangle(0, 0, chart1.Width, chart1.Height));
            });
            RectangleF bounds = e.PageSettings.PrintableArea;
            float factor = ((float)bounds.Height / (float)bmp.Width);
            e.Graphics.DrawImage(bmp, bounds.Left, 100, (factor * bmp.Width), (factor * bmp.Height));
        }

        private void button2_Click(object sender, EventArgs e)
        {


        }

        private void doc_PrintPage2(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Bitmap bmp = new Bitmap(label1.Width, label1.Height, label1.CreateGraphics());
            this.Invoke((MethodInvoker)delegate()
            {
                label1.DrawToBitmap(bmp, new Rectangle(0, 0, label1.Width, label1.Height));
            });
            RectangleF bounds = e.PageSettings.PrintableArea;
            float factor = ((float)bounds.Height / (float)bmp.Height);
            e.Graphics.DrawImage(bmp, bounds.Left + 100, bounds.Top, (factor * bmp.Width), (factor * bmp.Height));
        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintDialog MyPrintDialog = new PrintDialog();
            if (MyPrintDialog.ShowDialog() == DialogResult.OK)
            {
                // do on a helper thread...
                ThreadPool.QueueUserWorkItem(s =>
                {
                    System.Drawing.Printing.PrintDocument doc = new System.Drawing.Printing.PrintDocument();
                    doc.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(doc_PrintPage);
                    doc.DefaultPageSettings.Landscape = true;
                    doc.Print();
                });
            }
        }

        private void chart1_Click(object sender, EventArgs e)
        {
            MouseEventArgs inClick = (MouseEventArgs)e;

            if (inClick.Button == MouseButtons.Right)
            {
                contextMenuStripGraphPrint.Show(Cursor.Position);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            MouseEventArgs inClick = (MouseEventArgs)e;

            if (inClick.Button == MouseButtons.Right)
            {
                contextMenuStripTextPrint.Show(Cursor.Position);
            }
        }

        private void toolStripMenuItem25_Click(object sender, EventArgs e)
        {
            PrintDialog MyPrintDialog = new PrintDialog();
            if (MyPrintDialog.ShowDialog() == DialogResult.OK)
            {
                // do on a helper thread...
                ThreadPool.QueueUserWorkItem(s =>
                {
                    System.Drawing.Printing.PrintDocument doc = new System.Drawing.Printing.PrintDocument();
                    doc.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(doc_PrintPage2);
                    doc.DefaultPageSettings.Landscape = false;
                    doc.Print();
                });
            }
        }

        private void toolStripMenuItem26_Click(object sender, EventArgs e)
        {
            rtbIncoming.Text = "";
        }

        private void rtbIncoming_Click(object sender, EventArgs e)
        {
            MouseEventArgs inClick = (MouseEventArgs) e;

            if (inClick.Button == MouseButtons.Right)
            {
                contextMenuStripClear.Show(Cursor.Position);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;

            foreach (Form frm in fc)
            {
                if (frm is MasterFillerInterface)
                {
                    if (frm.WindowState == FormWindowState.Minimized)
                    {
                        frm.WindowState = FormWindowState.Normal;
                    }
                    frm.BringToFront();
                    return;
                }
            }
            MasterFillerInterface f2 = new MasterFillerInterface(dataGridView1.CurrentRow.Index,d.Rows[dataGridView1.CurrentRow.Index][1].ToString());
            f2.Owner = this;
            f2.Show();
        }

        private void cMSTestType_Opening(object sender, CancelEventArgs e)
        {

        }

        private void Main_Form_Shown(object sender, EventArgs e)
        {
            //finnally reformat the slave row so that cols 2,5,8 are lightsteelblue..
            for (int i = 0; i < 16; i++)
            {
                if (d.Rows[i][9].ToString().Contains("S"))
                {
                    dataGridView1.Rows[i].Cells[2].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[i].Cells[5].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[i].Cells[8].Style.BackColor = Color.LightSteelBlue;
                    dataGridView1.Rows[i].Cells[12].Style.BackColor = Color.LightSteelBlue;
                }
            }

            dataGridView1_Resize(this, null);
        }


    }// end mainform class section...

}
