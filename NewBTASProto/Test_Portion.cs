﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;
using System.Data.OleDb;
using System.Diagnostics;

namespace NewBTASProto
{
    public partial class Main_Form : Form
    {

        string comboText;

        public CancellationTokenSource[] cRunTest = new CancellationTokenSource[16];

        public static readonly object dataBaseLock = new object();

        //vars for recording in sc

        private void RunTest(string testType = "default", int station = -1)
        {

            ///
            /// General Structure:
            /// 
            /// As Recieved - Y - >  Jump straight to the test (case 1)
            ///  |
            ///  N
            ///  |
            ///  V
            ///  Intelligent Charger with Auto Config - Y - >  Do everything! (case 2)
            ///  |
            ///  N
            ///  |
            ///  V
            ///  Intelligent Charger - Y - >  Skip the configure part... (case 3)
            ///  |
            ///  N
            ///  |
            ///  V
            ///  Legacy Charger - Y - >  Limited version of the test (case 4)
            ///  |
            ///  N
            ///  |
            ///  V
            ///  Shunt - Y - >  Start test when you see a current (case 5)

            if (station == -1)
            {
                station = dataGridView1.CurrentRow.Index;
            }

            int Cstation = 0;

            //here to prevent double starts...
            if (cRunTest[station] != null && cRunTest[station].IsCancellationRequested == false)
            {
                return;
            }

            if (testType == "default")
            {
                testType = d.Rows[station][2].ToString();
            }

            #region master slave test and setup
            // we will use this bool to say if we need to do slave stuff..
            bool MasterSlaveTest = false;
            int slaveRow = -1;

            if (d.Rows[station][9].ToString() == "") { ;}  // do nothing if there is no assigned charger id
            else if (d.Rows[station][9].ToString().Length > 2)  // this is the case where we have a master and slave config
            {

                MasterSlaveTest = true;
                // split into 3 and 4 digit case
                if (d.Rows[station][9].ToString().Length == 3)
                {
                    // 3 case
                    Cstation = int.Parse(d.Rows[station][9].ToString().Substring(0, 1));
                }
                else
                {
                    // 4 case
                    Cstation = int.Parse(d.Rows[station][9].ToString().Substring(0, 2));
                }

                // also assign the slave channel...
                string temp = d.Rows[station][9].ToString().Replace("-M", "");

                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                    {
                        //found the slave
                        slaveRow = i;
                        break;
                    }
                }
            }  // end the master slave find...
            else  // this is the normal case with just one charger
            {
                Cstation = Convert.ToInt32(d.Rows[station][9]);
            }
            #endregion

            cRunTest[station] = new CancellationTokenSource();

            #region workorder number setup
            // we need to split up the work orders if we have multiple work orders on a single line...
            string tempWOS = d.Rows[station][1].ToString();
            char[] delims = { ' ' };
            string[] A = tempWOS.Split(delims);
            string MWO1 = A[0];
            string MWO2 = "";
            string MWO3 = "";
            if (A.Length > 2) { MWO2 = A[1]; }
            if (A.Length > 3) { MWO3 = A[2]; }

            string SWO1 = "";
            string SWO2 = "";
            string SWO3 = "";
            if (slaveRow != -1)
            {
                tempWOS = d.Rows[slaveRow][1].ToString();
                A = tempWOS.Split(delims);
                SWO1 = A[0];
                if (A.Length > 2) { SWO2 = A[1]; }
                if (A.Length > 3) { SWO3 = A[2]; }
            }
            #endregion

            // Everything is going to be done on a helper thread
            ThreadPool.QueueUserWorkItem(s =>
            {
                // master try / catch for test...
                try
                {

                    /////////////////////////////////SET UP/////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    #region SETUP

                    string strAccessConn;
                    string strAccessSelect;
                    OleDbConnection myAccessConn;

                    // setup the canellation token
                    CancellationToken token = (CancellationToken)s;

                    //cable connected test variables
                    bool shuntCon = false;
                    bool tempCon = false;
                    bool cellCon = false;
                    bool oldShuntCon = false;
                    bool oldTempCon = false;
                    bool oldCellCon = false;

                    byte badCSCanReads = 0;

                    double[,] cellHistory = new double[24, 4] { { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 } };
                    double[,] slaveCellHistory = new double[24, 4] { { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 }, { 0, 0, 0, 0 } };
                    DateTime cellCheckTimer = DateTime.Now;

                    bool runAsShunt = false;

                    //make sure we aren't running a shorting test in a master/slave config
                    if (d.Rows[station][2].ToString().Contains("Shorting") && MasterSlaveTest == true)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Cannot run a master/slave shorting test!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        cRunTest[station].Cancel();
                        return;
                    }
                    //also make sure we have the shorting cable connected when running a shorting test
                    if (d.Rows[station][2].ToString().Contains("Shorting") && !(GlobalVars.CScanData[station].CCID == 23 || GlobalVars.CScanData[station].CCID == 24))
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Cannot run shorting test without the shorting cable!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        cRunTest[station].Cancel();
                        return;
                    }


                    //do these checks for a normal test and the first test of a combo test only...
                    if (!d.Rows[station][2].ToString().Contains("Combo") || (d.Rows[station][2].ToString().Contains("Combo") && d.Rows[station][2].ToString().Contains("(1 ")))
                    {
                        #region startup checks (all cases)
                        // first we check if we have all the relavent options selected

                        if ((string)d.Rows[station][1] == "")
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "Please Assign a Work Order", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            cRunTest[station].Cancel();
                            return;
                        }
                        else if (testType == "")
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "Please Select a Test Type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            cRunTest[station].Cancel();
                            return;
                        }
                        else if ((bool)d.Rows[station][4] == false)
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "CScan is not In Use. Please Select it Before Proceeding.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            cRunTest[station].Cancel();
                            return;
                        }
                        else if (dataGridView1.Rows[station].Cells[4].Style.BackColor != Color.Green && dataGridView1.Rows[station].Cells[4].Style.BackColor != Color.YellowGreen)
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "CScan is not currently connected.  Please Check Connection.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            cRunTest[station].Cancel();
                            return;
                        }
                        else if (GlobalVars.CScanData[station].cellCableType == "NONE")
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "CScan is not connected to a cells Cable.  Please connect a cells cable to this CSCAN to run a test.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            cRunTest[station].Cancel();
                            return;
                        }
                        else if (GlobalVars.CScanData[station].shuntCableType == "NONE" && !(testType == "As Received" || testType.Contains("Short")))
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "CScan is not connected to a Shunt cable.  Please connect a shunt cable to this CSCAN to run a test.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            cRunTest[station].Cancel();
                            return;
                        }
                        // check if we have the right cells cable for multiple work orders on one cscan...
                        else if (MWO2 != "" && MWO3 == "" && GlobalVars.CScanData[station].CCID != 3)
                        {
                            // we need to make sure the master has a 2x11 to continue
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "You have two work orders assocaited with the master CScan, but do not have a 2X11 cable connected to it.  In order to record two work orders with one CScan you must use a 2X11 cable.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            cRunTest[station].Cancel();
                            return;
                        }
                        else if (MWO2 != "" && MWO3 != "" && GlobalVars.CScanData[station].CCID != 4)
                        {
                            // we need to make sure the master has a 3x7 to continue
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "You have three work orders assocaited with the master CScan, but do not have a 3X7 cable connected to it.  In order to record three work orders with one CScan you must use a 3X7 cable.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            cRunTest[station].Cancel();
                            return;
                        }
                        else if (MWO2 == "" && GlobalVars.CScanData[station].CCID == 3)
                        {
                            // warn the user that if they use a 2X11 cable with only one work order that the data associated with the second battery will be lost
                            this.Invoke((MethodInvoker)delegate()
                            {
                                DialogResult dialogResult = MessageBox.Show(this, "Are you sure you want to continue? You have one work order assocaited with the master CScan, but have a 2X11 cable connected to it.  In order to record the data from all 22 channels you will need to add an additional workorder to this station.", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                this.Invoke((MethodInvoker)delegate() { this.dataGridView1.Enabled = true; });
                                if (dialogResult == DialogResult.No)
                                {
                                    cRunTest[station].Cancel();
                                    return;
                                }
                            });
                            if (cRunTest[station].IsCancellationRequested == true) { return; }
                        }
                        else if ((MWO2 == "" || MWO3 == "") && GlobalVars.CScanData[station].CCID == 4)
                        {
                            // warn the user that if they use a 3x7 cable with only one or two work orders that the data associated with the second or third battery will be lost
                            this.Invoke((MethodInvoker)delegate()
                            {
                                DialogResult dialogResult = MessageBox.Show(this, "Are you sure you want to continue? You have less than 3 work orders assocaited with the master CScan, but have a 3X7 cable connected to it.  In order to record the data from all 21 channels you will need to add an additional workorder(s) to this station.", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                this.Invoke((MethodInvoker)delegate() { this.dataGridView1.Enabled = true; });
                                if (dialogResult == DialogResult.No)
                                {
                                    cRunTest[station].Cancel();
                                    return;
                                }
                            });
                            if (cRunTest[station].IsCancellationRequested == true) { return; }
                        }


                        //2x11 case
                        if (GlobalVars.CScanData[station].CCID == 3 && (int)pci.Rows[station][3] != (GlobalVars.CScanData[station].cellsToDisplay) && (int)pci.Rows[station][3] != -1)
                        {
                            // warn the user that the number of cells set in the database does not match the work order...
                            this.Invoke((MethodInvoker)delegate()
                            {
                                DialogResult dialogResult = MessageBox.Show(this, "The number of cells the battery contains does not match the cells cable currently being used.  Do you want to continue?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                                if (dialogResult == DialogResult.No)
                                {
                                    cRunTest[station].Cancel();
                                    return;
                                }
                            });
                            if (cRunTest[station].IsCancellationRequested == true) { return; }
                        }
                        //3X7 case
                        else if (GlobalVars.CScanData[station].CCID == 4 && (int)pci.Rows[station][3] != (GlobalVars.CScanData[station].cellsToDisplay) && (int)pci.Rows[station][3] != -1)
                        {
                            // warn the user that the number of cells set in the database does not match the work order...
                            this.Invoke((MethodInvoker)delegate()
                            {
                                DialogResult dialogResult = MessageBox.Show(this, "The number of cells the battery contains does not match the cells cable currently being used.  Do you want to continue?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dialogResult == DialogResult.No)
                                {
                                    cRunTest[station].Cancel();
                                    return;
                                }
                            });
                            if (cRunTest[station].IsCancellationRequested == true) { return; }
                        }
                        //general case
                        else if (pci.Rows[station][1].ToString().Contains("NiCd") && (int)pci.Rows[station][3] != GlobalVars.CScanData[station].cellsToDisplay && (int)pci.Rows[station][3] != -1)
                        {
                            // warn the user that the number of cells set in the database does not match the work order...
                            this.Invoke((MethodInvoker)delegate()
                            {
                                DialogResult dialogResult = MessageBox.Show(this, "The number of cells the battery contains does not match the cells cable currently being used.  Do you want to continue?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dialogResult == DialogResult.No)
                                {
                                    cRunTest[station].Cancel();
                                    return;
                                }
                            });
                            if (cRunTest[station].IsCancellationRequested == true) { return; }
                        }

                        // may also need to check some of this for the slave
                        if (MasterSlaveTest)
                        {
                            if ((string)d.Rows[slaveRow][1] == "")
                            {
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    MessageBox.Show(this, "Please Assign a Work Order to the Slave Channel", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                });
                                cRunTest[station].Cancel();
                                return;
                            }
                            else if ((bool)d.Rows[slaveRow][4] == false)
                            {
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    MessageBox.Show(this, "The slave CScan is not In Use. Please make sure that is is In Use and connected before proceeding.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                });
                                cRunTest[station].Cancel();
                                return;
                            }
                            else if (dataGridView1.Rows[slaveRow].Cells[4].Style.BackColor != Color.Green && dataGridView1.Rows[slaveRow].Cells[4].Style.BackColor != Color.YellowGreen)
                            {
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    MessageBox.Show(this, "Slave CScan is not currently connected.  Please Check Connection.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                });
                                cRunTest[station].Cancel();
                                return;
                            }
                            else if (GlobalVars.CScanData[slaveRow].cellCableType == "NONE")
                            {
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    MessageBox.Show(this, "Slave CScan is not connected to a cells Cable.  Please connect a cells cable to this CSCAN to run a test with it.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                });
                                cRunTest[station].Cancel();
                                return;
                            }
                            // check if we have the right cells cable for multiple work orders on one cscan...
                            else if (SWO2 != "" && SWO3 == "" && GlobalVars.CScanData[slaveRow].CCID != 3)
                            {
                                // we need to make sure the master has a 2x11 to continue
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    MessageBox.Show(this, "You have two work orders assocaited with the slave CScan, but do not have a 2X11 cable connected to it.  In order to record two work orders with one CScan you must use a 2X11 cable.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                });
                                cRunTest[station].Cancel();
                                return;
                            }
                            else if (SWO2 != "" && SWO3 != "" && GlobalVars.CScanData[slaveRow].CCID != 4)
                            {
                                // we need to make sure the master has a 2x11 to continue
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    MessageBox.Show(this, "You have three work orders assocaited with the slave CScan, but do not have a 3X7 cable connected to it.  In order to record three work orders with one CScan you must use a 3X7 cable.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                });
                                cRunTest[station].Cancel();
                                return;
                            }
                            else if (SWO2 == "" && GlobalVars.CScanData[slaveRow].CCID == 3)
                            {
                                // warn the user that if they use a 2X11 cable with only one work order that the data associated with the second battery will be lost
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    DialogResult dialogResult = MessageBox.Show(this, "Are you sure you want to continue? You have one work order assocaited with the slave CScan, but have a 2X11 cable connected to it.  In order to record the data from all 22 channels you will need to add an additional workorder to this station.", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                    if (dialogResult == DialogResult.No)
                                    {
                                        cRunTest[station].Cancel();
                                        return;
                                    }
                                });
                                if (cRunTest[station].IsCancellationRequested == true) { return; }

                            }
                            else if ((SWO2 == "" || SWO3 == "") && GlobalVars.CScanData[slaveRow].CCID == 4)
                            {
                                // warn the user that if they use a 3x7 cable with only one or two work orders that the data associated with the second or third battery will be lost
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    DialogResult dialogResult = MessageBox.Show(this, "Are you sure you want to continue? You have less than 3 work orders assocaited with the slave CScan, but have a 3X7 cable connected to it.  In order to record the data from all 21 channels you will need to add an additional workorder(s) to this station.", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                    if (dialogResult == DialogResult.No)
                                    {
                                        cRunTest[station].Cancel();
                                        return;
                                    }
                                });
                                if (cRunTest[station].IsCancellationRequested == true) { return; }
                            }

                            //2x11 case
                            if (GlobalVars.CScanData[slaveRow].CCID == 3 && (int)pci.Rows[slaveRow][3] != (GlobalVars.CScanData[slaveRow].cellsToDisplay) && (int)pci.Rows[station][3] != -1)
                            {
                                // warn the user that the number of cells set in the database does not match the work order...
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    DialogResult dialogResult = MessageBox.Show(this, "The number of cells the battery connected to the slave CSCAN contains does not match the cells cable currently being used.  Do you want to continue?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                    if (dialogResult == DialogResult.No)
                                    {
                                        cRunTest[station].Cancel();
                                        return;
                                    }
                                });
                                if (cRunTest[station].IsCancellationRequested == true) { return; }
                            }
                            //3X7 case
                            if (GlobalVars.CScanData[slaveRow].CCID == 3 && (int)pci.Rows[slaveRow][3] != (GlobalVars.CScanData[slaveRow].cellsToDisplay) && (int)pci.Rows[station][3] != -1)
                            {
                                // warn the user that the number of cells set in the database does not match the work order...
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    DialogResult dialogResult = MessageBox.Show(this, "The number of cells the battery connected to the slave CSCAN contains does not match the cells cable currently being used.  Do you want to continue?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                    if (dialogResult == DialogResult.No)
                                    {
                                        cRunTest[station].Cancel();
                                        return;
                                    }
                                });
                                if (cRunTest[station].IsCancellationRequested == true) { return; }
                            }
                            //general case
                            else if (pci.Rows[slaveRow][1].ToString().Contains("NiCd") && (int)pci.Rows[slaveRow][3] != GlobalVars.CScanData[slaveRow].cellsToDisplay && (int)pci.Rows[station][3] != -1)
                            {
                                // warn the user that the number of cells set in the database does not match the work order...
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    DialogResult dialogResult = MessageBox.Show(this, "The number of cells the battery connected to the slave CSCAN contains does not match the cells cable currently being used.  Do you want to continue?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                    if (dialogResult == DialogResult.No)
                                    {
                                        cRunTest[station].Cancel();
                                        return;
                                    }
                                });
                                if (cRunTest[station].IsCancellationRequested == true) { return; }
                            }

                            // check that the batteries are the same model!
                            // need to create a db connection
                            try
                            {
                                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                                myAccessConn = new OleDbConnection(strAccessConn);
                            }
                            catch (Exception ex)
                            {
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                });
                                cRunTest[station].Cancel();
                                return;
                            }


                            // Now i need to pull in the model values for both work orders...
                            // Need to think about expanding this to the possible 6 batteries....
                            try
                            {
                                // get the battery serial model
                                strAccessSelect = @"SELECT BatteryModel FROM WorkOrders WHERE WorkOrderNumber='" + MWO1 + "';";
                                DataSet batMod1 = new DataSet();
                                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                                lock (dataBaseLock)
                                {
                                    myAccessConn.Open();
                                    myDataAdapter.Fill(batMod1, "workOrder");
                                    myAccessConn.Close();
                                }
                                // get the battery serial model
                                strAccessSelect = @"SELECT BatteryModel FROM WorkOrders WHERE WorkOrderNumber='" + SWO1 + "';";
                                DataSet batMod2 = new DataSet();
                                myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                                myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                                lock (dataBaseLock)
                                {
                                    myAccessConn.Open();
                                    myDataAdapter.Fill(batMod2, "workOrder");
                                    myAccessConn.Close();
                                }

                                // we got both so let's compare!
                                if (batMod1.Tables[0].Rows[0][0].ToString() != batMod2.Tables[0].Rows[0][0].ToString())
                                {
                                    if (GlobalVars.autoConfig && (bool)d.Rows[station][12] && (d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")))
                                    {
                                        //we are setup for AutoConfig, but we have different types of batteries in the master slave, which will not work
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            MessageBox.Show(this, "Cannot AutoConfig the charger if the Master and Slave batteries are not the same.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        });
                                        cRunTest[station].Cancel();
                                        return;
                                    }
                                    else
                                    {
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            DialogResult dialogResult = MessageBox.Show(this, "The Master and Slave battery models do not match. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                            if (dialogResult == DialogResult.No)
                                            {
                                                cRunTest[station].Cancel();
                                                return;
                                            }
                                        });
                                        if (cRunTest[station].IsCancellationRequested == true) { return; }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    MessageBox.Show(this, "Error: Failed to pull data in from the Database. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                });
                                cRunTest[station].Cancel();
                                return;
                            }
                        }

                        if (d.Rows[station][10].ToString().Contains("Shunt"))
                        {
                            runAsShunt = true;
                        }
                        else if ((bool)d.Rows[station][8] == false && !(testType == "As Received" || testType.ToString().Contains("Shorting")))
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                DialogResult dialogResult = MessageBox.Show(this, "The charger is not linked. Do you want to proceed with the test in shunt mode?", "Click Yes to continue or No to cancel the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dialogResult == DialogResult.No)
                                {
                                    cRunTest[station].Cancel();
                                    return;
                                }
                            });
                            if (cRunTest[station].IsCancellationRequested == true) { return; }
                            runAsShunt = true;
                        }


                        if ((string)d.Rows[station][11] == "offline!" && !(testType == "As Received" || testType.ToString().Contains("Shorting")) && (d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) && !runAsShunt)
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "The Intelligent Charger is set to be offline.  Please set the charger to be online by pressing the following key sequence on the charger: FUNC, 1, 1 and ENTER.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            cRunTest[station].Cancel();
                            return;
                        }

                        else if ((d.Rows[station][9].ToString() == "" || d.Rows[station][10].ToString() == "") && !(testType == "As Received" || testType.ToString().Contains("Shorting")))
                        {
                            // we don't have a charger linked. Do we still want to continue...
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "There is no charger ID entered.  You must enter a charger ID to run any test other than 'As Received' or 'Shorting'", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            cRunTest[station].Cancel();
                            return;
                        }


                        //Finally Lets check if the charger is running also...
                        else if (d.Rows[station][11].ToString() == "RUN" && (d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) && !runAsShunt)
                        {
                            //looks like that charger is already running.  Lets ask the user if we should stop the charger or not.
                            this.Invoke((MethodInvoker)delegate()
                            {
                                DialogResult dialogResult = MessageBox.Show(this, "The charger appears to already be running. Do you want to stop it now and proceed with the test?", "Click Yes to have the program stop the charger or No to do it manually", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dialogResult == DialogResult.No)
                                {
                                    cRunTest[station].Cancel();
                                    return;
                                }
                            });
                            if (cRunTest[station].IsCancellationRequested == true) { return; }

                            // since we are still here...
                            // let's stop the charger...
                            updateD(station, 5, true);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, true); }
                            // and update the status column
                            updateD(station, 7, "Stopping Charger");
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Stopping Charger"); }

                            for (int j = 0; j < 5; j++)
                            {
                                // set KE1 to 2 ("command")
                                GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                                // set KE3 to stop
                                GlobalVars.ICSettings[Cstation].KE3 = (byte)2;
                                //Update the output string value
                                GlobalVars.ICSettings[Cstation].UpdateOutText();
                                criticalNum[Cstation] = true;
                                //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                                Thread.Sleep(5000);

                                // set KE1 to 1 ("query")
                                GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                                //Update the output string value
                                GlobalVars.ICSettings[Cstation].UpdateOutText();
                                for (int i = 0; i < 5; i++)
                                {
                                    criticalNum[Cstation] = true;
                                    Thread.Sleep(1000);
                                    if (d.Rows[station][11].ToString() == "HOLD")
                                    {
                                        break;
                                    }

                                }
                                if (d.Rows[station][11].ToString() == "HOLD")
                                {
                                    break;
                                }
                            }

                        }


                        cellCon = true;
                        oldCellCon = true;
                        if (GlobalVars.CScanData[station].shuntCableType != "NONE")
                        {
                            shuntCon = true;
                            oldShuntCon = true;
                        }
                        if (GlobalVars.CScanData[station].tempPlateType != "NONE")
                        {
                            tempCon = true;
                            oldTempCon = true;
                        }


                        #endregion
                    }
                    // we are running the second test of a combo...
                    else
                    {
                        if (d.Rows[station][10].ToString().Contains("Shunt"))
                        {
                            runAsShunt = true;
                        }
                        else if ((bool)d.Rows[station][8] == false && testType != "As Received")
                        {
                            runAsShunt = true;
                        }
                    }

                    #region db connection setup (all cases)
                    // create db connection
                    try
                    {
                        strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                        myAccessConn = new OleDbConnection(strAccessConn);
                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        cRunTest[station].Cancel();
                        return;
                    }

                    #endregion



                    // we passed the tests so we'll check the box to indicate we are a go!
                    updateD(station, 5, true);
                    if (MasterSlaveTest) { updateD(slaveRow, 5, true); }

                    #region test startup wait (only for the ICAs...)
                    if ((d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) && !runAsShunt)
                    {
                        if (readTLock() == true)
                        {
                            // we need to wait
                            //let the user know and then poll tlock for a while...
                            //update the GUI
                            updateD(station, 7, "Test Start Wait");
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Test Start Wait"); }

                            while (readTLock() == true)
                            {
                                Thread.Sleep(200);
                                //also look for a cancel...
                                if (token.IsCancellationRequested)
                                {


                                    //clear values from d
                                    updateD(station, 7, ("Cancelled"));
                                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Cancelled"); }
                                    updateD(station, 5, false);
                                    if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                                    //update the gui
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        sendNote(station, 2, "Test Cancelled");
                                    });
                                    cRunTest[station].Cancel();
                                    return;
                                }

                            }
                        }

                        //now we need to set tlock so that only one test starts at a time...

                        setTLock();
                        if (d.Rows[station][7].ToString() == "Test Start Wait")
                        {
                            Thread.Sleep(5000);
                        }

                    }


                    #endregion


                    #region load test readings and interval values (true for all cases)
                    // Now we'll load the test parameters
                    // We need to know the Interval and the number of readings///////////////////////////////////////////////////////////////////////

                    int readings;
                    int interval;

                    //  open the db and pull in the options table
                    try
                    {
                        strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME='" + testType + "';";
                        DataSet settings = new DataSet();
                        OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                        lock (dataBaseLock)
                        {
                            myAccessConn.Open();
                            myDataAdapter.Fill(settings, "TestType");
                            myAccessConn.Close();
                        }

                        readings = (int)settings.Tables[0].Rows[0][3];
                        interval = (int)settings.Tables[0].Rows[0][4];

                    }
                    catch (Exception ex)
                    {
                        myAccessConn.Close();

                        updateD(station, 5, false);
                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                        //clear values from d
                        updateD(station, 7, ("Error"));
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                        clearTLock();
                        cRunTest[station].Cancel();
                        this.Invoke((MethodInvoker)delegate()
                        {
                            sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        return;
                    }
                    #endregion

                    #region if we are doing the autoconfig, let's get the charger settings in order and then loaded into the charger! (case 2 only)

                    // these will be used later in the IC check portion to confirm that the charger is running as expected
                    float curSet1 = 0;
                    float vSet1 = 0;
                    float curSet2 = 0;
                    float vSet2 = 0;

                    // we'll tell the charger what to do! (if we have an IC and the user wants us to...)
                    if (GlobalVars.autoConfig && (bool)d.Rows[station][12] && (d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) && testType != "As Received" && !runAsShunt && !testType.Contains("Shorting"))
                    {

                        // GENERAL PROCEDURE
                        // We are going to look up the settings
                        // Tell the User to confirm the settings (later only let them directly change them
                        // Then load them into the charger

                        // first we need to pull in the settings from the DB
                        //  open the db and pull in the options table
                        try
                        {
                            // get the battery serial model
                            strAccessSelect = @"SELECT * FROM WorkOrders WHERE WorkOrderNumber='" + MWO1 + "';";
                            DataSet workOrder = new DataSet();
                            OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                            lock (dataBaseLock)
                            {
                                myAccessConn.Open();
                                myDataAdapter.Fill(workOrder, "workOrder");
                                myAccessConn.Close();
                            }

                            string model = (string)workOrder.Tables[0].Rows[0][4];

                            //now that we have the model we need to pull in the settings to load into the charger
                            strAccessSelect = @"SELECT * FROM BatteriesCustom WHERE BatteryModel='" + model + "';";
                            DataSet battery = new DataSet();
                            myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                            myDataAdapter = new OleDbDataAdapter(myAccessCommand);


                            lock (dataBaseLock)
                            {
                                myAccessConn.Open();
                                myDataAdapter.Fill(battery, "battery");
                                myAccessConn.Close();
                            }

                            //we will also pick up the test settings in case this is a custom test...
                            strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME='" + testType + "';";
                            DataSet testSettings = new DataSet();
                            myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                            myDataAdapter = new OleDbDataAdapter(myAccessCommand);


                            lock (dataBaseLock)
                            {
                                myAccessConn.Open();
                                myDataAdapter.Fill(testSettings, "testSettings");
                                myAccessConn.Close();
                            }

                            

                            // now we can assign the battery settings to the GlobalVars.
                            // We will decide on the settings based on the test being performed...
                            byte[] tempKMStore = new byte[21] { 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48 }; // 21 values for the 21 KM params

                            #region setting switch

                            try
                            {
                                switch (testType)
                                {
                                    case "Full Charge-6":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][42].ToString().Substring(0, 2)));          //mode
                                        tempKMStore[1] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][43].ToString()));                          //time hours
                                        tempKMStore[2] = Convert.ToByte(48);                                                                            //time mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + Math.Floor(GetDouble(battery.Tables[0].Rows[0][45].ToString()) * 10));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][45].ToString()) * 1000) % 100);                                                                            //bottom current byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + Math.Floor(GetDouble(battery.Tables[0].Rows[0][45].ToString()) / 1));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][45].ToString()) % 1) * 100);                                                                            //bottom current byte
                                        }
                                        else
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + Math.Floor(GetDouble(battery.Tables[0].Rows[0][45].ToString()) / 10));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (10 * (GetDouble(battery.Tables[0].Rows[0][45].ToString()) % 10)));        //bottom current byte
                                        }
                                        tempKMStore[5] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][46].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][46].ToString()))) / 1));              //top voltage byte
                                        tempKMStore[6] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * GetDouble(battery.Tables[0].Rows[0][46].ToString())) : GetDouble(battery.Tables[0].Rows[0][46].ToString())) % 1));            //bottom voltage byte

                                        tempKMStore[7] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][47].ToString()));                          //time 2 hours
                                        tempKMStore[8] = (GlobalVars.AddOneMin ? Convert.ToByte(49) : Convert.ToByte(48));                                  //time 2 mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[9] = Convert.ToByte(48 + Math.Floor(GetDouble(battery.Tables[0].Rows[0][49].ToString())) * 10);             //top current 2 byte
                                            tempKMStore[10] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][49].ToString()) * 1000) % 100);                                                                           //bottom current 2 byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[9] = Convert.ToByte(48 + Math.Floor(GetDouble(battery.Tables[0].Rows[0][49].ToString())) / 1);             //top current 2 byte
                                            tempKMStore[10] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][49].ToString()) % 1) * 100);                                                                           //bottom current 2 byte
                                        }
                                        else
                                        {
                                            tempKMStore[9] = Convert.ToByte(48 + Math.Floor(GetDouble(battery.Tables[0].Rows[0][49].ToString()) / 10));             //top current 2 byte
                                            tempKMStore[10] = Convert.ToByte(48 + (10 * (GetDouble(battery.Tables[0].Rows[0][49].ToString()) % 10)));       //bottom current 2 byte
                                        }
                                        tempKMStore[11] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][50].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][50].ToString()))) / 1));                 //top voltage 2 byte
                                        tempKMStore[12] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * GetDouble(battery.Tables[0].Rows[0][50].ToString())) : GetDouble(battery.Tables[0].Rows[0][50].ToString())) % 1));           //bottom voltage 2 byte

                                        tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                        tempKMStore[14] = (byte)48;                                                                                //discharge time mins
                                        tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                        tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                        tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                        tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                        tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                        tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                        curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][45].ToString());
                                        vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][46].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][46].ToString()));
                                        curSet2 = (float)GetDouble(battery.Tables[0].Rows[0][49].ToString());
                                        vSet2 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][50].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][50].ToString()));
                                        break;
                                    case "Full Charge-4":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][52].ToString().Substring(0, 2)));          //mode
                                        tempKMStore[1] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][53].ToString()));                          //time hours
                                        //tempKMStore[2] = (byte)(48);                                                                            //time mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][55].ToString())) * 10));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][55].ToString()) * 1000) % 100);                                                                            //bottom current byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][55].ToString())) / 1));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][55].ToString()) % 1) * 100);                                                                            //bottom current byte
                                        }
                                        else
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][55].ToString()) / 10)));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (10 * (GetDouble(battery.Tables[0].Rows[0][55].ToString()) % 10)));        //bottom current byte
                                        }
                                        tempKMStore[5] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][56].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][56].ToString()))) / 1));                  //top voltage byte
                                        tempKMStore[6] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * GetDouble(battery.Tables[0].Rows[0][56].ToString())) : GetDouble(battery.Tables[0].Rows[0][56].ToString())) % 1));            //bottom current byte

                                        tempKMStore[7] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][57].ToString()));                          //time 2 hours
                                        tempKMStore[8] = (GlobalVars.AddOneMin ? Convert.ToByte(49) : Convert.ToByte(48));                                  //time 2 mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][59].ToString()) * 10)));             //top current 2 byte
                                            tempKMStore[10] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][59].ToString()) * 1000) % 100);                                                                           //bottom current 2 byte
                                        }
                                        if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][59].ToString()) / 1)));             //top current 2 byte
                                            tempKMStore[10] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][59].ToString()) % 1) * 100);                                                                           //bottom current 2 byte
                                        }
                                        else
                                        {
                                            tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][59].ToString())) / 10));             //top current 2 byte
                                            tempKMStore[10] = Convert.ToByte(48 + (10 * (GetDouble(battery.Tables[0].Rows[0][59].ToString()) % 10)));       //bottom current 2 byte
                                        }
                                        tempKMStore[11] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][60].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][60].ToString()))) / 1));                 //top voltage 2 byte
                                        tempKMStore[12] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * GetDouble(battery.Tables[0].Rows[0][60].ToString())) : GetDouble(battery.Tables[0].Rows[0][60].ToString())) % 1));           //bottom voltage 2 byte

                                        tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                        tempKMStore[14] = (byte)48;                                                                              //discharge time mins
                                        tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                        tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                        tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                        tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                        tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                        tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                        curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][55].ToString());
                                        vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][56].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][56].ToString()));
                                        curSet2 = (float)GetDouble(battery.Tables[0].Rows[0][59].ToString());
                                        vSet2 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][60].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][60].ToString()));
                                        break;
                                    case "Top Charge-4":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][62].ToString().Substring(0, 2)));          //mode
                                        tempKMStore[1] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][63].ToString()));                          //time hours
                                        tempKMStore[2] = (GlobalVars.AddOneMin ? Convert.ToByte(49) : Convert.ToByte(48));                                  //time mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][65].ToString())) * 10));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][65].ToString()) * 1000) % 100);                                                                            //bottom current byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][65].ToString())) / 1));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][65].ToString()) % 1) * 100);                                                                            //bottom current byte
                                        }
                                        else
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][65].ToString()) / 10)));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (10 * (GetDouble(battery.Tables[0].Rows[0][65].ToString()) % 10)));        //bottom current byte
                                        }
                                        tempKMStore[5] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][66].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][66].ToString()))) / 1));              //top voltage byte
                                        tempKMStore[6] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * GetDouble(battery.Tables[0].Rows[0][66].ToString())) : GetDouble(battery.Tables[0].Rows[0][66].ToString())) % 1));            //bottom current byte

                                        tempKMStore[7] = (byte)48;                                                                                 //time 2 hours
                                        tempKMStore[8] = (byte)48;                                                                                 //time 2 mins
                                        tempKMStore[9] = (byte)48;                                                                                 //top current 2 byte
                                        tempKMStore[10] = (byte)48;                                                                                //bottom current 2 byte
                                        tempKMStore[11] = (byte)48;                                                                                //top voltage 2 byte
                                        tempKMStore[12] = (byte)48;                                                                                //bottom voltage 2 byte

                                        tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                        tempKMStore[14] = (byte)48;                                                                                //discharge time mins
                                        tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                        tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                        tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                        tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                        tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                        tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                        curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][65].ToString());
                                        vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][66].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][66].ToString()));
                                        curSet2 = 0;
                                        vSet2 = 0;
                                        break;
                                    case "Top Charge-2":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][72].ToString().Substring(0, 2)));          //mode
                                        tempKMStore[1] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][73].ToString()));                          //time hours
                                        tempKMStore[2] = (GlobalVars.AddOneMin ? Convert.ToByte(49) : Convert.ToByte(48));                                  //time mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][75].ToString())) * 10));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][75].ToString()) * 1000) % 100);                                                                            //bottom current byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][75].ToString())) / 1));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][75].ToString()) % 1) * 100);                                                                            //bottom current byte
                                        }
                                        else
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][75].ToString()) / 10)));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (10 * (GetDouble(battery.Tables[0].Rows[0][75].ToString()) % 10)));        //bottom current byte
                                        }
                                        tempKMStore[5] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][76].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][76].ToString()))) / 1));                  //top voltage byte
                                        tempKMStore[6] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * GetDouble(battery.Tables[0].Rows[0][76].ToString())) : GetDouble(battery.Tables[0].Rows[0][76].ToString())) % 1));            //bottom current byte

                                        tempKMStore[7] = (byte)48;                                                                                 //time 2 hours
                                        tempKMStore[8] = (byte)48;                                                                                 //time 2 mins
                                        tempKMStore[9] = (byte)48;                                                                                 //top current 2 byte
                                        tempKMStore[10] = (byte)48;                                                                                //bottom current 2 byte
                                        tempKMStore[11] = (byte)48;                                                                                //top voltage 2 byte
                                        tempKMStore[12] = (byte)48;                                                                                //bottom voltage 2 byte

                                        tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                        tempKMStore[14] = (byte)48;                                                                                //discharge time mins
                                        tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                        tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                        tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                        tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                        tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                        tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                        curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][75].ToString());
                                        vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][76].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][76].ToString()));
                                        curSet2 = 0;
                                        vSet2 = 0;
                                        break;
                                    case "Top Charge-1":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][82].ToString().Substring(0, 2)));          //mode
                                        tempKMStore[1] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][83].ToString()));                          //time hours
                                        tempKMStore[2] = (GlobalVars.AddOneMin ? Convert.ToByte(49) : Convert.ToByte(48));                                  //time mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][85].ToString())) * 10));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][85].ToString()) * 1000) % 100);                                                                            //bottom current byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][85].ToString())) / 1));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][85].ToString()) % 1) * 100);                                                                            //bottom current byte
                                        }
                                        else
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][85].ToString()) / 10)));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][85].ToString())) % 10));        //bottom current byte
                                        }
                                        tempKMStore[5] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][86].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][86].ToString()))) / 1));                  //top voltage byte
                                        tempKMStore[6] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(battery.Tables[0].Rows[0][86].ToString()))) : (GetDouble(battery.Tables[0].Rows[0][86].ToString()))) % 1));            //bottom current byte

                                        tempKMStore[7] = (byte)48;                                                                                 //time 2 hours
                                        tempKMStore[8] = (byte)48;                                                                                 //time 2 mins
                                        tempKMStore[9] = (byte)48;                                                                                 //top current 2 byte
                                        tempKMStore[10] = (byte)48;                                                                                //bottom current 2 byte
                                        tempKMStore[11] = (byte)48;                                                                                //top voltage 2 byte
                                        tempKMStore[12] = (byte)48;                                                                                //bottom voltage 2 byte

                                        tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                        tempKMStore[14] = (byte)48;                                                                                //discharge time mins
                                        tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                        tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                        tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                        tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                        tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                        tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                        curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][85].ToString());
                                        vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][86].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][86].ToString()));
                                        curSet2 = 0;
                                        vSet2 = 0;
                                        break;
                                    case "Capacity-1":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][92].ToString().Substring(0, 2)));          //mode
                                        tempKMStore[1] = (byte)48;                                                                                 //time hours
                                        tempKMStore[2] = (byte)48;                                                                                 //time mins
                                        tempKMStore[3] = (byte)48;                                                                                 //top current byte
                                        tempKMStore[4] = (byte)48;                                                                                 //bottom current byte
                                        tempKMStore[5] = (byte)48;                                                                                 //top voltage byte
                                        tempKMStore[6] = (byte)48;                                                                                 //bottom current byte

                                        tempKMStore[7] = (byte)48;                                                                                 //time 2 hours
                                        tempKMStore[8] = (byte)48;                                                                                 //time 2 mins
                                        tempKMStore[9] = (byte)48;                                                                                 //top current 2 byte
                                        tempKMStore[10] = (byte)48;                                                                                //bottom current 2 byte
                                        tempKMStore[11] = (byte)48;                                                                                //top voltage 2 byte
                                        tempKMStore[12] = (byte)48;                                                                                //bottom voltage 2 byte

                                        tempKMStore[13] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][93].ToString()));                         //discharge time hours
                                        tempKMStore[14] = (GlobalVars.AddOneMin ? Convert.ToByte(49) : Convert.ToByte(48));                                 //discharge time mins
                                        if (tempKMStore[0] == 31 + 48)
                                        {
                                            if (d.Rows[station][10].ToString().Contains("mini"))
                                            {
                                                tempKMStore[15] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][95].ToString())) * 1));         //discharge current high byte
                                                tempKMStore[16] = Convert.ToByte(48 + 100 * (GetDouble(battery.Tables[0].Rows[0][95].ToString())) % 1);   //discharge current low byte
                                            }
                                            else
                                            {
                                                tempKMStore[15] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][95].ToString()) / 10)));        //discharge current high byte
                                                tempKMStore[16] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][95].ToString())) % 10));   //discharge current low byte
                                            }
                                        }
                                        tempKMStore[17] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][96].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][96].ToString()))) / 1));                 //discharge voltage high byte
                                        tempKMStore[18] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * GetDouble(battery.Tables[0].Rows[0][96].ToString())) : (GetDouble(battery.Tables[0].Rows[0][96].ToString()))) % 1));           //discharge voltage low byte
                                        if (tempKMStore[0] == 32 + 48)
                                        {
                                            tempKMStore[19] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][101].ToString())) / 1));            //discharge resistance high byte
                                            tempKMStore[20] = Convert.ToByte(48 + 100 * (GetDouble(battery.Tables[0].Rows[0][101].ToString())) % 1);      //discharge resistance low byte
                                        }
                                        if (tempKMStore[0] == 31 + 48) { curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][95].ToString()); }
                                        vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][96].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][96].ToString()));
                                        curSet2 = 0;
                                        vSet2 = 0;
                                        break;
                                    case "Discharge":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][102].ToString().Substring(0, 2)));         //mode
                                        tempKMStore[1] = (byte)48;                                                                                 //time hours
                                        tempKMStore[2] = (byte)48;                                                                                 //time mins
                                        tempKMStore[3] = (byte)48;                                                                                 //top current byte
                                        tempKMStore[4] = (byte)48;                                                                                 //bottom current byte
                                        tempKMStore[5] = (byte)48;                                                                                 //top voltage byte
                                        tempKMStore[6] = (byte)48;                                                                                 //bottom current byte

                                        tempKMStore[7] = (byte)48;                                                                                 //time 2 hours
                                        tempKMStore[8] = (byte)48;                                                                                 //time 2 mins
                                        tempKMStore[9] = (byte)48;                                                                                 //top current 2 byte
                                        tempKMStore[10] = (byte)48;                                                                                //bottom current 2 byte
                                        tempKMStore[11] = (byte)48;                                                                                //top voltage 2 byte
                                        tempKMStore[12] = (byte)48;                                                                                //bottom voltage 2 byte

                                        tempKMStore[13] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][103].ToString()));                        //discharge time hours
                                        tempKMStore[14] = (GlobalVars.AddOneMin ? Convert.ToByte(49) : Convert.ToByte(48));                                 //discharge time mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[15] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][105].ToString())) * 1));            //discharge current high byte
                                            tempKMStore[16] = Convert.ToByte(48 + (100 * ((float)GetDouble(battery.Tables[0].Rows[0][105].ToString())) % 1));      //discharge current low byte
                                        }
                                        else
                                        {
                                            tempKMStore[15] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][105].ToString()) / 10)));           //discharge current high byte
                                            tempKMStore[16] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][105].ToString())) % 10));      //discharge current low byte
                                        }
                                        tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                        tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                        tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                        tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                        curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][105].ToString());
                                        vSet1 = 0;
                                        curSet2 = 0;
                                        vSet2 = 0;
                                        break;
                                    case "Slow Charge-14":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][112].ToString().Substring(0, 2)));         //mode
                                        tempKMStore[1] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][113].ToString()));                         //time hours
                                        tempKMStore[2] = (GlobalVars.AddOneMin ? Convert.ToByte(49) : Convert.ToByte(48));                                  //time mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][115].ToString()) * 10)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][115].ToString()) * 1000) % 100);                                                                            //bottom current byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][115].ToString()) / 1)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][115].ToString()) % 1) * 100);                                                                            //bottom current byte
                                        }
                                        else
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][115].ToString()) / 10)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][115].ToString())) % 10));       //bottom current byte
                                        }
                                        tempKMStore[5] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][116].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][116].ToString()))) / 1));                 //top voltage byte
                                        tempKMStore[6] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(battery.Tables[0].Rows[0][116].ToString()))) : (GetDouble(battery.Tables[0].Rows[0][116].ToString()))) % 1));           //bottom current byte

                                        tempKMStore[7] = (byte)48;                                                                                 //time 2 hours
                                        tempKMStore[8] = (byte)48;                                                                                 //time 2 mins
                                        tempKMStore[9] = (byte)48;                                                                                 //top current 2 byte
                                        tempKMStore[10] = (byte)48;                                                                                //bottom current 2 byte
                                        tempKMStore[11] = (byte)48;                                                                                //top voltage 2 byte
                                        tempKMStore[12] = (byte)48;                                                                                //bottom voltage 2 byte

                                        tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                        tempKMStore[14] = (byte)48;                                                                                //discharge time mins
                                        tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                        tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                        tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                        tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                        tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                        tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                        curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][115].ToString());
                                        vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][116].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][116].ToString()));
                                        curSet2 = 0;
                                        vSet2 = 0;
                                        break;
                                    case "Slow Charge-16":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][122].ToString().Substring(0, 2)));         //mode
                                        tempKMStore[1] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][123].ToString()));                         //time hours
                                        tempKMStore[2] = (GlobalVars.AddOneMin ? Convert.ToByte(49) : Convert.ToByte(48));                                  //time mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][125].ToString())) * 10));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][125].ToString()) * 1000) % 100);                                                                            //bottom current byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][125].ToString())) / 1));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][125].ToString()) % 1) * 100);                                                                            //bottom current byte
                                        }
                                        else
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][125].ToString()) / 10)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][125].ToString())) % 10));       //bottom current byte
                                        }
                                        tempKMStore[5] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][126].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][126].ToString()))) / 1));                 //top voltage byte
                                        tempKMStore[6] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(battery.Tables[0].Rows[0][126].ToString()))) : (GetDouble(battery.Tables[0].Rows[0][126].ToString()))) % 1));           //bottom current byte

                                        tempKMStore[7] = (byte)48;                                                                                 //time 2 hours
                                        tempKMStore[8] = (byte)48;                                                                                 //time 2 mins
                                        tempKMStore[9] = (byte)48;                                                                                 //top current 2 byte
                                        tempKMStore[10] = (byte)48;                                                                                //bottom current 2 byte
                                        tempKMStore[11] = (byte)48;                                                                                //top voltage 2 byte
                                        tempKMStore[12] = (byte)48;                                                                                //bottom voltage 2 byte

                                        tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                        tempKMStore[14] = (byte)48;                                                                                //discharge time mins
                                        tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                        tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                        tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                        tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                        tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                        tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                        curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][125].ToString());
                                        vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][126].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][126].ToString()));
                                        curSet2 = 0;
                                        vSet2 = 0;
                                        break;
                                    case "Custom Chg":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][132].ToString().Substring(0, 2)));             //mode
                                        tempKMStore[1] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][7].ToString()));                          //time hours
                                        tempKMStore[2] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][8].ToString()));                          //time mins

                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][135].ToString()) * 10)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][135].ToString()) * 1000) % 100);                                                                            //bottom current byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][135].ToString()) / 1)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][135].ToString()) % 1) * 100);                                                                            //bottom current byte
                                        }
                                        else
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][135].ToString()) / 10)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][135].ToString())) % 10));       //bottom current byte
                                        }
                                        tempKMStore[5] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][136].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][136].ToString()))) / 1));                 //top voltage byte
                                        tempKMStore[6] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(battery.Tables[0].Rows[0][136].ToString()))) : (GetDouble(battery.Tables[0].Rows[0][136].ToString()))) % 1));           //bottom current byte

                                        if (tempKMStore[0] == (20 + 48) || tempKMStore[0] == (21 + 48))
                                        {
                                            tempKMStore[7] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][11].ToString()));                          //time 2 hours
                                            tempKMStore[8] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][12].ToString()));                          //time 2 mins
                                            if (d.Rows[station][10].ToString().Contains("mini"))
                                            {
                                                tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][139].ToString())) * 10));        //top current 2 byte
                                                tempKMStore[10] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][139].ToString()) * 1000) % 100);                                                                       //bottom current 2 byte
                                            }
                                            else if (d.Rows[station][10].ToString().Contains("MFC"))
                                            {
                                                tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][139].ToString())) / 1));        //top current 2 byte
                                                tempKMStore[10] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][139].ToString()) % 1) * 100);                                                                       //bottom current 2 byte
                                            }
                                            else
                                            {
                                                tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][139].ToString()) / 10)));        //top current 2 byte
                                                tempKMStore[10] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][139].ToString())) % 10));  //bottom current 2 byte
                                            }
                                            tempKMStore[11] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][140].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][140].ToString()))) / 1));            //top voltage 2 byte
                                            tempKMStore[12] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(battery.Tables[0].Rows[0][140].ToString()))) : (GetDouble(battery.Tables[0].Rows[0][140].ToString()))) % 1));      //bottom voltage 2 byte
                                        }

                                        tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                        tempKMStore[14] = (byte)48;                                                                                //discharge time mins
                                        tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                        tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                        tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                        tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                        tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                        tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                        curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][135].ToString());
                                        vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][136].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][136].ToString()));
                                        if (battery.Tables[0].Rows[0][139].ToString() != "") { curSet2 = (float)GetDouble(battery.Tables[0].Rows[0][139].ToString()); }
                                        if (battery.Tables[0].Rows[0][140].ToString() != "") { vSet2 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][140].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][140].ToString())); }
                                        break;
                                    case "Custom Cap":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][142].ToString().Substring(0, 2)));         //mode
                                        tempKMStore[1] = (byte)48;                                                                                 //time hours
                                        tempKMStore[2] = (byte)48;                                                                                 //time mins
                                        tempKMStore[3] = (byte)48;                                                                                 //top current byte
                                        tempKMStore[4] = (byte)48;                                                                                 //bottom current byte
                                        tempKMStore[5] = (byte)48;                                                                                 //top voltage byte
                                        tempKMStore[6] = (byte)48;                                                                                 //bottom current byte

                                        tempKMStore[7] = (byte)48;                                                                                 //time 2 hours
                                        tempKMStore[8] = (byte)48;                                                                                 //time 2 mins
                                        tempKMStore[9] = (byte)48;                                                                                 //top current 2 byte
                                        tempKMStore[10] = (byte)48;                                                                                //bottom current 2 byte
                                        tempKMStore[11] = (byte)48;                                                                                //top voltage 2 byte
                                        tempKMStore[12] = (byte)48;                                                                                //bottom voltage 2 byte

                                        tempKMStore[13] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][16].ToString()));                         //discharge time hours
                                        tempKMStore[14] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][17].ToString()));                         //discharge time mins
                                        if (tempKMStore[0] != 32 + 48)
                                        {
                                            if (d.Rows[station][10].ToString().Contains("mini"))
                                            {
                                                tempKMStore[15] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][145].ToString())) * 1));        //discharge current high byte
                                                tempKMStore[16] = Convert.ToByte(48 + (100 * ((GetDouble(battery.Tables[0].Rows[0][145].ToString())) % 1)));  //discharge current low byte
                                            }
                                            else
                                            {
                                                tempKMStore[15] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][145].ToString()) / 10)));       //discharge current high byte
                                                tempKMStore[16] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][145].ToString())) % 10));  //discharge current low byte
                                            }
                                        }
                                        if (tempKMStore[0] != 30 + 48)
                                        {
                                            tempKMStore[17] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][146].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][146].ToString()))) / 1));            //discharge voltage high byte
                                            tempKMStore[18] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(battery.Tables[0].Rows[0][146].ToString()))) : (GetDouble(battery.Tables[0].Rows[0][146].ToString()))) % 1));      //discharge voltage low byte
                                        }
                                        if (tempKMStore[0] == 32 + 48)
                                        {
                                            tempKMStore[19] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][151].ToString())) / 1));            //discharge resistance high byte
                                            tempKMStore[20] = Convert.ToByte(48 + 100 * ((GetDouble(battery.Tables[0].Rows[0][151].ToString())) % 1));      //discharge resistance low byte
                                        }
                                        if (battery.Tables[0].Rows[0][145].ToString() != "") { curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][145].ToString()); }
                                        if (battery.Tables[0].Rows[0][146].ToString() != "") { vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][146].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][146].ToString())); }
                                        curSet2 = 0;
                                        vSet2 = 0;
                                        break;
                                    case "Constant Voltage":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][152].ToString().Substring(0, 2)));         //mode
                                        tempKMStore[1] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][153].ToString()));                         //time hours
                                        tempKMStore[2] = (GlobalVars.AddOneMin ? Convert.ToByte(49) : Convert.ToByte(48));                                  //time mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][155].ToString())) * 10));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][155].ToString()) * 1000) % 100);                                                                            //bottom current byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][155].ToString())) / 1));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][155].ToString()) % 1) * 100);                                                                            //bottom current byte
                                        }
                                        else
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][155].ToString()) / 10)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][155].ToString())) % 10));       //bottom current byte
                                        }
                                        tempKMStore[5] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][156].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][156].ToString()))) / 1));                 //top voltage byte
                                        tempKMStore[6] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(battery.Tables[0].Rows[0][156].ToString()))) : (GetDouble(battery.Tables[0].Rows[0][156].ToString()))) % 1));           //bottom voltage byte

                                        tempKMStore[7] = (byte)48;                                                                                 //time 2 hours
                                        tempKMStore[8] = (byte)48;                                                                                 //time 2 mins
                                        tempKMStore[9] = (byte)48;                                                                                 //top current 2 byte
                                        tempKMStore[10] = (byte)48;                                                                                //bottom current 2 byte
                                        tempKMStore[11] = (byte)48;                                                                                //top voltage 2 byte
                                        tempKMStore[12] = (byte)48;                                                                                //bottom voltage 2 byte

                                        tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                        tempKMStore[14] = (byte)48;                                                                                //discharge time mins
                                        tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                        tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                        tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                        tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                        tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                        tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                        curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][155].ToString());
                                        vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][156].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][156].ToString()));
                                        curSet2 = 0;
                                        vSet2 = 0;
                                        break;
                                    case "Full Charge-4.5":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][162].ToString().Substring(0, 2)));          //mode
                                        tempKMStore[1] = (byte)(48 + 2);                          //time hours
                                        tempKMStore[2] = (byte)(48 + 30);                                                                            //time mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][165].ToString())) * 10));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][165].ToString()) * 1000) % 100);                                                                            //bottom current byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][165].ToString())) / 1));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][165].ToString()) % 1) * 100);                                                                            //bottom current byte
                                        }
                                        else
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][165].ToString()) / 10)));             //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][165].ToString())) % 10));        //bottom current byte
                                        }
                                        tempKMStore[5] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][166].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][166].ToString()))) / 1));                  //top voltage byte
                                        tempKMStore[6] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(battery.Tables[0].Rows[0][166].ToString()))) : (GetDouble(battery.Tables[0].Rows[0][166].ToString()))) % 1));            //bottom current byte

                                        tempKMStore[7] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][167].ToString()));                          //time 2 hours
                                        tempKMStore[8] = (GlobalVars.AddOneMin ? Convert.ToByte(49) : Convert.ToByte(48));                                   //time 2 mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][169].ToString())) * 10));             //top current 2 byte
                                            tempKMStore[10] = Convert.ToByte((48 + (GetDouble(battery.Tables[0].Rows[0][169].ToString()) * 1000) % 100));                                                                           //bottom current 2 byte
                                        }
                                        if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][169].ToString())) / 1));             //top current 2 byte
                                            tempKMStore[10] = Convert.ToByte((48 + (GetDouble(battery.Tables[0].Rows[0][169].ToString()) % 1) * 100));                                                                           //bottom current 2 byte
                                        }
                                        else
                                        {
                                            tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][169].ToString()) / 10)));             //top current 2 byte
                                            tempKMStore[10] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][169].ToString())) % 10));       //bottom current 2 byte
                                        }
                                        tempKMStore[11] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][170].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][170].ToString()))) / 1));                 //top voltage 2 byte
                                        tempKMStore[12] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(battery.Tables[0].Rows[0][170].ToString()))) : (GetDouble(battery.Tables[0].Rows[0][170].ToString()))) % 1));           //bottom voltage 2 byte

                                        tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                        tempKMStore[14] = (byte)48;                                                                              //discharge time mins
                                        tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                        tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                        tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                        tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                        tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                        tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                        curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][165].ToString());
                                        vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][166].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][166].ToString()));
                                        curSet2 = (float)GetDouble(battery.Tables[0].Rows[0][169].ToString());
                                        vSet2 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][170].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][170].ToString()));
                                        break;
                                    case "Custom Chg 2":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][172].ToString().Substring(0, 2)));         //mode
                                        tempKMStore[1] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][7].ToString()));                          //time hours
                                        tempKMStore[2] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][8].ToString()));                          //time mins

                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][175].ToString()) * 10)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][175].ToString()) * 1000) % 100);                                                                            //bottom current byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][175].ToString()) / 1)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][175].ToString()) % 1) * 100);                                                                            //bottom current byte
                                        }
                                        else
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][175].ToString()) / 10)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][175].ToString())) % 10));       //bottom current byte
                                        }
                                        tempKMStore[5] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][176].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][176].ToString()))) / 1));                 //top voltage byte
                                        tempKMStore[6] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(battery.Tables[0].Rows[0][176].ToString()))) : (GetDouble(battery.Tables[0].Rows[0][176].ToString()))) % 1));           //bottom current byte

                                        if (tempKMStore[0] == (20 + 48) || tempKMStore[0] == (21 + 48))
                                        {
                                            tempKMStore[7] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][11].ToString()));                          //time 2 hours
                                            tempKMStore[8] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][12].ToString()));                          //time 2 mins
                                            if (d.Rows[station][10].ToString().Contains("mini"))
                                            {
                                                tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][179].ToString())) * 10));        //top current 2 byte
                                                tempKMStore[10] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][179].ToString()) * 1000) % 100);                                                                       //bottom current 2 byte
                                            }
                                            else if (d.Rows[station][10].ToString().Contains("MFC"))
                                            {
                                                tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][179].ToString())) / 1));        //top current 2 byte
                                                tempKMStore[10] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][179].ToString()) % 1) * 100);                                                                       //bottom current 2 byte
                                            }
                                            else
                                            {
                                                tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][179].ToString()) / 10)));        //top current 2 byte
                                                tempKMStore[10] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][179].ToString())) % 10));  //bottom current 2 byte
                                            }
                                            tempKMStore[11] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][180].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][180].ToString()))) / 1));            //top voltage 2 byte
                                            tempKMStore[12] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(battery.Tables[0].Rows[0][180].ToString()))) : (GetDouble(battery.Tables[0].Rows[0][180].ToString()))) % 1));      //bottom voltage 2 byte
                                        }

                                        tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                        tempKMStore[14] = (byte)48;                                                                                //discharge time mins
                                        tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                        tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                        tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                        tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                        tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                        tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                        curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][175].ToString());
                                        vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][176].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][176].ToString()));
                                        if (battery.Tables[0].Rows[0][179].ToString() != "") { curSet2 = (float)GetDouble(battery.Tables[0].Rows[0][179].ToString()); }
                                        if (battery.Tables[0].Rows[0][180].ToString() != "") { vSet2 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][180].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][180].ToString())); }
                                        break;
                                    case "Custom Chg 3":
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(battery.Tables[0].Rows[0][182].ToString().Substring(0, 2)));         //mode
                                        tempKMStore[1] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][7].ToString()));                          //time hours
                                        tempKMStore[2] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][8].ToString()));                          //time mins

                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][185].ToString()) * 10)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][185].ToString()) * 1000) % 100);                                                                            //bottom current byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][185].ToString()) / 1)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][185].ToString()) % 1) * 100);                                                                            //bottom current byte
                                        }
                                        else
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][185].ToString()) / 10)));            //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][185].ToString())) % 10));       //bottom current byte
                                        }
                                        tempKMStore[5] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][186].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][186].ToString()))) / 1));                 //top voltage byte
                                        tempKMStore[6] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(battery.Tables[0].Rows[0][186].ToString()))) : (GetDouble(battery.Tables[0].Rows[0][186].ToString()))) % 1));           //bottom current byte

                                        if (tempKMStore[0] == (20 + 48) || tempKMStore[0] == (21 + 48))
                                        {
                                            tempKMStore[7] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][11].ToString()));                          //time 2 hours
                                            tempKMStore[8] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][12].ToString()));                          //time 2 mins
                                            if (d.Rows[station][10].ToString().Contains("mini"))
                                            {
                                                tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][189].ToString())) * 10));        //top current 2 byte
                                                tempKMStore[10] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][189].ToString()) * 1000) % 100);                                                                       //bottom current 2 byte
                                            }
                                            else if (d.Rows[station][10].ToString().Contains("MFC"))
                                            {
                                                tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][189].ToString())) / 1));        //top current 2 byte
                                                tempKMStore[10] = Convert.ToByte(48 + (GetDouble(battery.Tables[0].Rows[0][189].ToString()) % 1) * 100);                                                                       //bottom current 2 byte
                                            }
                                            else
                                            {
                                                tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(battery.Tables[0].Rows[0][189].ToString()) / 10)));        //top current 2 byte
                                                tempKMStore[10] = Convert.ToByte(48 + 10 * ((GetDouble(battery.Tables[0].Rows[0][189].ToString())) % 10));  //bottom current 2 byte
                                            }
                                            tempKMStore[11] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(battery.Tables[0].Rows[0][190].ToString()))) : Math.Floor(GetDouble(battery.Tables[0].Rows[0][190].ToString()))) / 1));            //top voltage 2 byte
                                            tempKMStore[12] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(battery.Tables[0].Rows[0][190].ToString()))) : (GetDouble(battery.Tables[0].Rows[0][190].ToString()))) % 1));      //bottom voltage 2 byte
                                        }

                                        tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                        tempKMStore[14] = (byte)48;                                                                                //discharge time mins
                                        tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                        tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                        tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                        tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                        tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                        tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                        curSet1 = (float)GetDouble(battery.Tables[0].Rows[0][185].ToString());
                                        vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][186].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][186].ToString()));
                                        if (battery.Tables[0].Rows[0][189].ToString() != "") { curSet2 = (float)GetDouble(battery.Tables[0].Rows[0][189].ToString()); }
                                        if (battery.Tables[0].Rows[0][190].ToString() != "") { vSet2 = (MasterSlaveTest ? (2 * (float)GetDouble(battery.Tables[0].Rows[0][190].ToString())) : (float)GetDouble(battery.Tables[0].Rows[0][190].ToString())); }
                                        break;
                                    default:
                                        // we got a custom test!
                                        //Lets load the custom test setting and try to load those then!
                                        //Everything is one to one in custom tests...  Much easier...
                                        tempKMStore[0] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][6].ToString().Substring(0, 2)));                                  //mode
                                        tempKMStore[1] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][7].ToString()));                                                  //time hours
                                        tempKMStore[2] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][8].ToString()));                                                  //time mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][9].ToString())) * 10));                                //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(testSettings.Tables[0].Rows[0][9].ToString()) * 1000) % 100);              //bottom current byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MC-10"))
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][9].ToString())) / 1));                                //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + (GetDouble(testSettings.Tables[0].Rows[0][9].ToString()) % 1) * 100);              //bottom current byte
                                        }
                                        else
                                        {
                                            tempKMStore[3] = Convert.ToByte(48 + (Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][9].ToString()) / 10)));                                //top current byte
                                            tempKMStore[4] = Convert.ToByte(48 + 10 * ((GetDouble(testSettings.Tables[0].Rows[0][9].ToString())) % 10));                           //bottom current byte
                                        }
                                        tempKMStore[5] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][10].ToString()))) : Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][10].ToString()))) / 1));                  //top voltage byte
                                        tempKMStore[6] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(testSettings.Tables[0].Rows[0][10].ToString()))) : (GetDouble(testSettings.Tables[0].Rows[0][10].ToString()))) % 1));            //bottom voltage byte

                                        tempKMStore[7] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][11].ToString()));                          //time 2 hours
                                        tempKMStore[8] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][12].ToString()));                          //time 2 mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][13].ToString())) * 10));             //top current 2 byte
                                            tempKMStore[10] = Convert.ToByte(48 + ((GetDouble(testSettings.Tables[0].Rows[0][13].ToString()) * 1000) % 100));                                                                           //bottom current 2 byte
                                        }
                                        else if (d.Rows[station][10].ToString().Contains("MFC"))
                                        {
                                            tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][13].ToString())) / 1));             //top current 2 byte
                                            tempKMStore[10] = Convert.ToByte(48 + ((GetDouble(testSettings.Tables[0].Rows[0][13].ToString()) % 1) * 100));                                                                           //bottom current 2 byte
                                        }
                                        else
                                        {
                                            tempKMStore[9] = Convert.ToByte(48 + (Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][13].ToString()) / 10)));             //top current 2 byte
                                            tempKMStore[10] = Convert.ToByte(48 + 10 * ((GetDouble(testSettings.Tables[0].Rows[0][13].ToString())) % 10));       //bottom current 2 byte
                                        }
                                        tempKMStore[11] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][14].ToString()))) : Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][14].ToString()))) / 1));                 //top voltage 2 byte
                                        tempKMStore[12] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(testSettings.Tables[0].Rows[0][14].ToString()))) : (GetDouble(testSettings.Tables[0].Rows[0][14].ToString()))) % 1));           //bottom voltage 2 byte

                                        tempKMStore[13] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][16].ToString()));                         //discharge time hours
                                        tempKMStore[14] = Convert.ToByte(48 + int.Parse(testSettings.Tables[0].Rows[0][17].ToString()));                         //discharge time mins
                                        if (tempKMStore[0] == 31 + 48 || tempKMStore[0] == 30 + 48)
                                        {
                                            if (d.Rows[station][10].ToString().Contains("mini"))
                                            {
                                                tempKMStore[15] = Convert.ToByte(48 + (Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][18].ToString())) * 1));         //discharge current high byte
                                                tempKMStore[16] = Convert.ToByte(48 + (100 * (GetDouble(testSettings.Tables[0].Rows[0][18].ToString()) % 1)));   //discharge current low byte
                                            }
                                            else
                                            {
                                                tempKMStore[15] = Convert.ToByte(48 + (Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][18].ToString()) / 10)));        //discharge current high byte
                                                tempKMStore[16] = Convert.ToByte(48 + 10 * ((GetDouble(testSettings.Tables[0].Rows[0][18].ToString())) % 10));   //discharge current low byte
                                            }
                                        }
                                        if (tempKMStore[0] == 31 + 48 || tempKMStore[0] == 32 + 48)
                                        {
                                            tempKMStore[17] = Convert.ToByte(48 + ((MasterSlaveTest ? (2 * Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][19].ToString()))) : Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][19].ToString()))) / 1));                 //discharge voltage high byte
                                            tempKMStore[18] = Convert.ToByte(48 + 100 * ((MasterSlaveTest ? (2 * (GetDouble(testSettings.Tables[0].Rows[0][19].ToString()))) : (GetDouble(testSettings.Tables[0].Rows[0][19].ToString()))) % 1));           //discharge voltage low byte
                                        }
                                        if (tempKMStore[0] == 32 + 48)
                                        {
                                            tempKMStore[19] = Convert.ToByte(48 + (Math.Floor(GetDouble(testSettings.Tables[0].Rows[0][15].ToString())) / 1));            //discharge resistance high byte
                                            tempKMStore[20] = Convert.ToByte(48 + 100 * ((GetDouble(testSettings.Tables[0].Rows[0][15].ToString())) % 1));      //discharge resistance low byte
                                        }

                                        // needs to be based on mode...
                                        if (tempKMStore[0] == 10 + 48 || tempKMStore[0] == 11 + 48 || tempKMStore[0] == 12 + 48)
                                        {
                                            curSet1 = (float)GetDouble(testSettings.Tables[0].Rows[0][9].ToString());
                                            vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(testSettings.Tables[0].Rows[0][10].ToString())) : (float)GetDouble(testSettings.Tables[0].Rows[0][10].ToString()));
                                            curSet2 = 0;
                                            vSet2 = 0;
                                        }
                                        else if (tempKMStore[0] == 20 + 48 || tempKMStore[0] == 21 + 48)
                                        {
                                            curSet1 = (float)GetDouble(testSettings.Tables[0].Rows[0][9].ToString());
                                            vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(testSettings.Tables[0].Rows[0][10].ToString())) : (float)GetDouble(testSettings.Tables[0].Rows[0][10].ToString()));
                                            curSet2 = (float)GetDouble(testSettings.Tables[0].Rows[0][13].ToString());
                                            vSet2 = (MasterSlaveTest ? (2 * (float)GetDouble(testSettings.Tables[0].Rows[0][14].ToString())) : (float)GetDouble(testSettings.Tables[0].Rows[0][14].ToString()));

                                        }
                                        else if (tempKMStore[0] == 30 + 48)
                                        {
                                            curSet1 = (float)GetDouble(testSettings.Tables[0].Rows[0][18].ToString());
                                            vSet1 = 0;
                                            curSet2 = 0;
                                            vSet2 = 0;
                                        }
                                        else if (tempKMStore[0] == 31 + 48 || tempKMStore[0] == 32 + 48)
                                        {
                                            if (tempKMStore[0] == 31 + 48) { curSet1 = (float)GetDouble(testSettings.Tables[0].Rows[0][18].ToString()); }
                                            vSet1 = (MasterSlaveTest ? (2 * (float)GetDouble(testSettings.Tables[0].Rows[0][19].ToString())) : (float)GetDouble(testSettings.Tables[0].Rows[0][19].ToString()));
                                            curSet2 = 0;
                                            vSet2 = 0;
                                        }
                                        break;
                                }// end switch
                            }
                            catch
                            {


                                //clear values from d
                                updateD(station, 7, ("Error"));
                                if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                                updateD(station, 5, false);
                                if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                                clearTLock();
                                cRunTest[station].Cancel();
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    sendNote(station, 2, "Error:  Fail to pull the settings from the DataBase.");
                                    if (testType == "Full Charge-4.5" ||
                                        testType == "Full Charge-4" ||
                                        testType == "Full Charge-6" ||
                                        testType == "Custom Cap" ||
                                        testType == "Constant Voltage" ||
                                        testType == "Custom Chg" ||
                                        testType == "Custom Chg 2" ||
                                        testType == "Custom Chg 3" ||
                                        testType == "Slow Charge-16" ||
                                        testType == "Slow Charge-14" ||
                                        testType == "Discharge" ||
                                        testType == "Capacity-1" ||
                                        testType == "Top Charge-1" ||
                                        testType == "Top Charge-2" ||
                                        testType == "Top Charge-4")
                                    {
                                        MessageBox.Show(this, "Failed to pull the settings from the DataBase. \r\nPlease make sure you have the battery model setup for this test under the Manage Battery Models menu.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                    else
                                    {
                                        MessageBox.Show(this, "Failed to pull the settings from the DataBase. \r\nPlease make sure you have the Test completely setup under the Edit Custom Test Settings menu.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                });

                                return;
                            }

                            #endregion

                            #region load settings
                            // set KE1 to data
                            GlobalVars.ICSettings[Cstation].KE1 = (byte)1;

                            // update KM1
                            GlobalVars.ICSettings[Cstation].KM1 = tempKMStore[0];
                            //// Charge Time 1
                            GlobalVars.ICSettings[Cstation].KM2 = tempKMStore[1];
                            GlobalVars.ICSettings[Cstation].KM3 = tempKMStore[2];
                            //// Charge Current 1
                            GlobalVars.ICSettings[Cstation].KM4 = tempKMStore[3];
                            GlobalVars.ICSettings[Cstation].KM5 = tempKMStore[4];
                            //// Charge Voltage 1
                            GlobalVars.ICSettings[Cstation].KM6 = tempKMStore[5];
                            GlobalVars.ICSettings[Cstation].KM7 = tempKMStore[6];


                            // Charge Time 2
                            GlobalVars.ICSettings[Cstation].KM8 = tempKMStore[7];
                            GlobalVars.ICSettings[Cstation].KM9 = tempKMStore[8];
                            // Charge Current 2
                            GlobalVars.ICSettings[Cstation].KM10 = tempKMStore[9];
                            GlobalVars.ICSettings[Cstation].KM11 = tempKMStore[10];
                            // Charge Voltage 2
                            GlobalVars.ICSettings[Cstation].KM12 = tempKMStore[11];
                            GlobalVars.ICSettings[Cstation].KM13 = tempKMStore[12];


                            // Discharge Time
                            GlobalVars.ICSettings[Cstation].KM14 = tempKMStore[13];
                            GlobalVars.ICSettings[Cstation].KM15 = tempKMStore[14];
                            // Discharge Current
                            GlobalVars.ICSettings[Cstation].KM16 = tempKMStore[15];
                            GlobalVars.ICSettings[Cstation].KM17 = tempKMStore[16];
                            // Discharge Voltage
                            GlobalVars.ICSettings[Cstation].KM18 = tempKMStore[17];
                            GlobalVars.ICSettings[Cstation].KM19 = tempKMStore[18];
                            // Discharge Resistance
                            GlobalVars.ICSettings[Cstation].KM20 = tempKMStore[19];
                            GlobalVars.ICSettings[Cstation].KM21 = tempKMStore[20];

                            //Update the output string value
                            GlobalVars.ICSettings[Cstation].UpdateOutText();

                            updateD(station, 7, "Loading Settings");
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Loading Settings"); }

                            criticalNum[Cstation] = true;
                            Thread.Sleep(2000);
                            criticalNum[Cstation] = true;
                            Thread.Sleep(2000);
                            criticalNum[Cstation] = true;
                            Thread.Sleep(2000);
                            criticalNum[Cstation] = true;
                            Thread.Sleep(2000);
                            // set KE1 to 0 ("data")
                            GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                            GlobalVars.ICSettings[Cstation].UpdateOutText();
                            criticalNum[Cstation] = true;
                            Thread.Sleep(2000);

                            #endregion

                        } // end try
                        catch (Exception ex)
                        {
                            myAccessConn.Close();

                            // reset everything
                            //clear values from d
                            updateD(station, 7, ("Error"));
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                            clearTLock();
                            cRunTest[station].Cancel();

                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 2, "Error:  Failed to Auto Load settings into the Charger.");
                                MessageBox.Show(this, "Error: Failed to Auto Load settings into the Charger.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            return;
                        } // end catch

                    }// end if

                    #endregion


                    #region check for new or resume and get step numbers
                    // step number variables for the 6 possible workorders..
                    int stepNum;
                    int stepNum2 = 0;
                    int stepNum3 = 0;
                    int slaveStepNum = 0;
                    int slaveStepNum2 = 0;
                    int slaveStepNum3 = 0;


                    //check if this is a new test first! Also check if the this is a master slave test and the two rows are out of sink...
                    if ((string)d.Rows[station][6] == "" || (MasterSlaveTest && d.Rows[station][6].ToString() != d.Rows[slaveRow][6].ToString()))
                    {

                        #region set up test number and ID
                        // Now we'll look up the current test number and increment the new step number/////////////////////////////////////////////////////


                        //  open the db and pull in the options table
                        try
                        {
                            strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + MWO1 + "' ORDER BY StepNumber DESC;";
                            DataSet tests = new DataSet();
                            OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                            lock (dataBaseLock)
                            {
                                myAccessConn.Open();
                                myDataAdapter.Fill(tests, "Tests");
                                myAccessConn.Close();
                            }

                            if (tests.Tables[0].Rows.Count == 0)
                            {
                                stepNum = 1;
                            }
                            else
                            {
                                stepNum = int.Parse((string)tests.Tables[0].Rows[0][4]) + 1;
                            }

                            if (MWO2 != "")
                            {
                                strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + MWO2 + "' ORDER BY StepNumber DESC;";
                                tests = new DataSet();
                                myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                                myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                                lock (dataBaseLock)
                                {
                                    myAccessConn.Open();
                                    myDataAdapter.Fill(tests, "Tests");
                                    myAccessConn.Close();
                                }

                                if (tests.Tables[0].Rows.Count == 0)
                                {
                                    stepNum2 = 1;
                                }
                                else
                                {
                                    stepNum2 = int.Parse((string)tests.Tables[0].Rows[0][4]) + 1;
                                }
                            }

                            if (MWO3 != "")
                            {
                                strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + MWO3 + "' ORDER BY StepNumber DESC;";
                                tests = new DataSet();
                                myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                                myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                                lock (dataBaseLock)
                                {
                                    myAccessConn.Open();
                                    myDataAdapter.Fill(tests, "Tests");
                                    myAccessConn.Close();
                                }

                                if (tests.Tables[0].Rows.Count == 0)
                                {
                                    stepNum3 = 1;
                                }
                                else
                                {
                                    stepNum3 = int.Parse((string)tests.Tables[0].Rows[0][4]) + 1;
                                }
                            }


                        }
                        catch (Exception ex)
                        {

                            myAccessConn.Close();
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                            //clear values from d
                            updateD(station, 7, ("Error"));
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                            clearTLock();
                            cRunTest[station].Cancel();
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            return;
                        }


                        //do we need to do the same for the slave?
                        if (MasterSlaveTest)
                        {
                            // Now we'll look up the current test number and increment the new step number/////////////////////////////////////////////////////
                            //  open the db and pull in the options table
                            try
                            {
                                strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + SWO1 + "' ORDER BY StepNumber DESC;";
                                DataSet tests = new DataSet();
                                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                                lock (dataBaseLock)
                                {
                                    myAccessConn.Open();
                                    myDataAdapter.Fill(tests, "Tests");
                                    myAccessConn.Close();
                                }

                                if (tests.Tables[0].Rows.Count == 0)
                                {
                                    slaveStepNum = 1;
                                }
                                else
                                {
                                    slaveStepNum = int.Parse((string)tests.Tables[0].Rows[0][4]) + 1;
                                }

                                if (SWO2 != "")
                                {
                                    strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + SWO2 + "' ORDER BY StepNumber DESC;";
                                    tests = new DataSet();
                                    myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                                    myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                                    lock (dataBaseLock)
                                    {
                                        myAccessConn.Open();
                                        myDataAdapter.Fill(tests, "Tests");
                                        myAccessConn.Close();
                                    }

                                    if (tests.Tables[0].Rows.Count == 0)
                                    {
                                        slaveStepNum2 = 1;
                                    }
                                    else
                                    {
                                        slaveStepNum2 = int.Parse((string)tests.Tables[0].Rows[0][4]) + 1;
                                    }
                                }

                                if (SWO3 != "")
                                {
                                    strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + SWO3 + "' ORDER BY StepNumber DESC;";
                                    tests = new DataSet();
                                    myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                                    myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                                    lock (dataBaseLock)
                                    {
                                        myAccessConn.Open();
                                        myDataAdapter.Fill(tests, "Tests");
                                        myAccessConn.Close();
                                    }

                                    if (tests.Tables[0].Rows.Count == 0)
                                    {
                                        slaveStepNum3 = 1;
                                    }
                                    else
                                    {
                                        slaveStepNum3 = int.Parse((string)tests.Tables[0].Rows[0][4]) + 1;
                                    }
                                }


                            }
                            catch (Exception ex)
                            {
                                myAccessConn.Close();

                                updateD(station, 5, false);
                                if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                //clear values from d
                                updateD(station, 7, ("Error"));
                                if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                                clearTLock();
                                cRunTest[station].Cancel();
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                });
                                return;
                            }
                        }




                        #endregion

                        #region save new test to test table
                        // Save the test information to the test table///////////////////////////////////////////////////////////////////////////////////////////

                        // we need the technician selected in the combo box too..
                        this.Invoke((MethodInvoker)delegate()
                        {
                            comboText = label2.Text;
                        });


                        //  now try to INSERT INTO it
                        try
                        {
                            string strUpdateCMD = "INSERT INTO Tests (WorkOrderId,WorkOrderNumber,AggrWorkOrders,StepNumber,[TestName],Readings,[Interval]," +
                                "StationNumber,Charger,DateStarted,DateCompleted,Technician,TerminalID,CellCableID,ShuntCableID,TempCableID,TerminalNumber," +
                                "BATNUMCABLE10) "
                                + "VALUES ('" +
                                "0" + "','" +                                                          //WorkOrderID (don't care..)
                                MWO1.Trim() + "','" +                         //WorkOrderNumber
                                "" + "','" +                                                           //AggrWorkOrders
                                stepNum.ToString("00") + "','" +                                       //StepNumber
                                testType + "','" +                                //TestName
                                readings.ToString() + "','" +                                          //Reading
                                (interval * 1000).ToString() + "','" +                                 //interval in msec
                                station.ToString() + "','" +                                           // station number
                                d.Rows[station][10].ToString() + "',#" +                               // charger type
                                DateTime.Now.Add(new TimeSpan(0, 0, 2)).ToString("yyyy-MM-dd HH:mm:ss") + "#,#" +                 // start date
                                DateTime.Now.Add(new TimeSpan(0, 0, 2)).ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                 // date completed
                                comboText + "','" +                                                    // technician
                                (GlobalVars.CScanData[station].terminalID + 216).ToString() + "','" +  //terminal ID
                                GlobalVars.CScanData[station].CCID.ToString() + "','" +                //cells cable ID
                                GlobalVars.CScanData[station].SHCID.ToString() + "','" +               //shunt cable ID
                                GlobalVars.CScanData[station].TCAB.ToString() + "','" +                //temp cable ID
                                d.Rows[station][9].ToString() + "','" +                                //charger ID (Terminal Number)
                                GlobalVars.CScanData[station].batNumCable10.ToString() +               //BATNUMCABLE10
                                "');";


                            OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                            lock (dataBaseLock)
                            {
                                myAccessConn.Open();
                                myAccessCommand.ExecuteNonQuery();
                                myAccessConn.Close();
                            }

                            if (MWO2 != "")
                            {
                                strUpdateCMD = "INSERT INTO Tests (WorkOrderId,WorkOrderNumber,AggrWorkOrders,StepNumber,[TestName],Readings,[Interval]," +
                                "StationNumber,Charger,DateStarted,DateCompleted,Technician,TerminalID,CellCableID,ShuntCableID,TempCableID,TerminalNumber," +
                                "BATNUMCABLE10) "
                                + "VALUES ('" +
                                "0" + "','" +                                                           //WorkOrderID (don't care..)
                                MWO2.Trim() + "','" +                                                   //WorkOrderNumber
                                "" + "','" +                                                            //AggrWorkOrders
                                stepNum2.ToString("00") + "','" +                                       //StepNumber
                                testType + "','" +                                                      //TestName
                                readings.ToString() + "','" +                                           //Reading
                                (interval * 1000).ToString() + "','" +                                  //interval in msec
                                station.ToString() + "','" +                                            // station number
                                d.Rows[station][10].ToString() + "',#" +                                // charger type
                                DateTime.Now.Add(new TimeSpan(0, 0, 2)).ToString("yyyy-MM-dd HH:mm:ss") + "#,#" +                  // start date
                                DateTime.Now.Add(new TimeSpan(0, 0, 2)).ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  // date completed
                                comboText + "','" +                                                     // technician
                                (GlobalVars.CScanData[station].terminalID + 216).ToString() + "','" +   //terminal ID
                                GlobalVars.CScanData[station].CCID.ToString() + "','" +                 //cells cable ID
                                GlobalVars.CScanData[station].SHCID.ToString() + "','" +                //shunt cable ID
                                GlobalVars.CScanData[station].TCAB.ToString() + "','" +                 //temp cable ID
                                d.Rows[station][9].ToString() + "','" +                                 //charger ID (Terminal Number)
                                GlobalVars.CScanData[station].batNumCable10.ToString() +                //BATNUMCABLE10
                                "');";


                                myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                                lock (dataBaseLock)
                                {
                                    myAccessConn.Open();
                                    myAccessCommand.ExecuteNonQuery();
                                    myAccessConn.Close();
                                }
                            }

                            if (MWO3 != "")
                            {
                                strUpdateCMD = "INSERT INTO Tests (WorkOrderId,WorkOrderNumber,AggrWorkOrders,StepNumber,[TestName],Readings,[Interval]," +
                                "StationNumber,Charger,DateStarted,DateCompleted,Technician,TerminalID,CellCableID,ShuntCableID,TempCableID,TerminalNumber," +
                                "BATNUMCABLE10) "
                                + "VALUES ('" +
                                "0" + "','" +                                                           //WorkOrderID (don't care..)
                                MWO3.Trim() + "','" +                                                   //WorkOrderNumber
                                "" + "','" +                                                            //AggrWorkOrders
                                stepNum3.ToString("00") + "','" +                                       //StepNumber
                                testType + "','" +                                                      //TestName
                                readings.ToString() + "','" +                                           //Reading
                                (interval * 1000).ToString() + "','" +                                  //interval in msec
                                station.ToString() + "','" +                                            // station number
                                d.Rows[station][10].ToString() + "',#" +                                // charger type
                                DateTime.Now.Add(new TimeSpan(0, 0, 2)).ToString("yyyy-MM-dd HH:mm:ss") + "#,#" +                  // start date
                                DateTime.Now.Add(new TimeSpan(0, 0, 2)).ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  // date completed
                                comboText + "','" +                                                     // technician
                                (GlobalVars.CScanData[station].terminalID + 216).ToString() + "','" +   //terminal ID
                                GlobalVars.CScanData[station].CCID.ToString() + "','" +                 //cells cable ID
                                GlobalVars.CScanData[station].SHCID.ToString() + "','" +                //shunt cable ID
                                GlobalVars.CScanData[station].TCAB.ToString() + "','" +                 //temp cable ID
                                d.Rows[station][9].ToString() + "','" +                                 //charger ID (Terminal Number)
                                GlobalVars.CScanData[station].batNumCable10.ToString() +                //BATNUMCABLE10
                                "');";


                                myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                                lock (dataBaseLock)
                                {
                                    myAccessConn.Open();
                                    myAccessCommand.ExecuteNonQuery();
                                    myAccessConn.Close();
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                            myAccessConn.Close();

                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                            //clear values from d
                            updateD(station, 7, ("Error"));
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                            clearTLock();
                            cRunTest[station].Cancel();
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            return;
                        }



                        //We made it to this point without errors, so we'll update the grid with the step number
                        updateD(station, 3, stepNum.ToString());

                        // We need to do the same with the slave (if there is one)
                        if (MasterSlaveTest)
                        {
                            try
                            {
                                string strUpdateCMD = "INSERT INTO Tests (WorkOrderId,WorkOrderNumber,AggrWorkOrders,StepNumber,[TestName],Readings,[Interval]," +
                                    "StationNumber,Charger,DateStarted,DateCompleted,Technician,TerminalID,CellCableID,ShuntCableID,TempCableID,TerminalNumber," +
                                    "BATNUMCABLE10) "
                                    + "VALUES ('" +
                                    "0" + "','" +                                                          //WorkOrderID (don't care..)
                                    SWO1.Trim() + "','" +                         //WorkOrderNumber
                                    "" + "','" +                                                           //AggrWorkOrders
                                    slaveStepNum.ToString("00") + "','" +                                       //StepNumber
                                    d.Rows[slaveRow][2].ToString() + "','" +                                //TestName
                                    readings.ToString() + "','" +                                          //Reading
                                    (interval * 1000).ToString() + "','" +                                 //interval in msec
                                    slaveRow.ToString() + "','" +                                           // slaveRow number
                                    d.Rows[slaveRow][10].ToString() + "',#" +                               // charger type
                                    DateTime.Now.Add(new TimeSpan(0, 0, 2)).ToString("yyyy-MM-dd HH:mm:ss") + "#,#" +                 // start date
                                    DateTime.Now.Add(new TimeSpan(0, 0, 2)).ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                 // date completed
                                    comboText + "','" +                                                    // technician
                                    (GlobalVars.CScanData[slaveRow].terminalID + 216).ToString() + "','" +  //terminal ID
                                    GlobalVars.CScanData[slaveRow].CCID.ToString() + "','" +                //cells cable ID
                                    GlobalVars.CScanData[slaveRow].SHCID.ToString() + "','" +               //shunt cable ID
                                    GlobalVars.CScanData[slaveRow].TCAB.ToString() + "','" +                //temp cable ID
                                    d.Rows[slaveRow][9].ToString() + "','" +                                //charger ID (Terminal Number)
                                    GlobalVars.CScanData[slaveRow].batNumCable10.ToString() +               //BATNUMCABLE10
                                    "');";


                                OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                                lock (dataBaseLock)
                                {
                                    myAccessConn.Open();
                                    myAccessCommand.ExecuteNonQuery();
                                    myAccessConn.Close();
                                }

                                if (SWO2 != "")
                                {
                                    strUpdateCMD = "INSERT INTO Tests (WorkOrderId,WorkOrderNumber,AggrWorkOrders,StepNumber,[TestName],Readings,[Interval]," +
                                    "StationNumber,Charger,DateStarted,DateCompleted,Technician,TerminalID,CellCableID,ShuntCableID,TempCableID,TerminalNumber," +
                                    "BATNUMCABLE10) "
                                    + "VALUES ('" +
                                    "0" + "','" +                                                           //WorkOrderID (don't care..)
                                    SWO2.Trim() + "','" +                                                   //WorkOrderNumber
                                    "" + "','" +                                                            //AggrWorkOrders
                                    slaveStepNum2.ToString("00") + "','" +                                  //StepNumber
                                    d.Rows[slaveRow][2].ToString() + "','" +                                //TestName
                                    readings.ToString() + "','" +                                           //Reading
                                    (interval * 1000).ToString() + "','" +                                  //interval in msec
                                    slaveRow.ToString() + "','" +                                           // slaveRow number
                                    d.Rows[slaveRow][10].ToString() + "',#" +                               // charger type
                                    DateTime.Now.Add(new TimeSpan(0, 0, 2)).ToString("yyyy-MM-dd HH:mm:ss") + "#,#" +                  // start date
                                    DateTime.Now.Add(new TimeSpan(0, 0, 2)).ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  // date completed
                                    comboText + "','" +                                                     // technician
                                    (GlobalVars.CScanData[slaveRow].terminalID + 216).ToString() + "','" +  //terminal ID
                                    GlobalVars.CScanData[slaveRow].CCID.ToString() + "','" +                //cells cable ID
                                    GlobalVars.CScanData[slaveRow].SHCID.ToString() + "','" +               //shunt cable ID
                                    GlobalVars.CScanData[slaveRow].TCAB.ToString() + "','" +                //temp cable ID
                                    d.Rows[slaveRow][9].ToString() + "','" +                                //charger ID (Terminal Number)
                                    GlobalVars.CScanData[slaveRow].batNumCable10.ToString() +               //BATNUMCABLE10
                                    "');";


                                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                                    lock (dataBaseLock)
                                    {
                                        myAccessConn.Open();
                                        myAccessCommand.ExecuteNonQuery();
                                        myAccessConn.Close();
                                    }
                                }

                                if (SWO3 != "")
                                {
                                    strUpdateCMD = "INSERT INTO Tests (WorkOrderId,WorkOrderNumber,AggrWorkOrders,StepNumber,[TestName],Readings,[Interval]," +
                                    "StationNumber,Charger,DateStarted,DateCompleted,Technician,TerminalID,CellCableID,ShuntCableID,TempCableID,TerminalNumber," +
                                    "BATNUMCABLE10) "
                                    + "VALUES ('" +
                                    "0" + "','" +                                                           //WorkOrderID (don't care..)
                                    SWO3.Trim() + "','" +                                                   //WorkOrderNumber
                                    "" + "','" +                                                            //AggrWorkOrders
                                    slaveStepNum3.ToString("00") + "','" +                                  //StepNumber
                                    d.Rows[slaveRow][2].ToString() + "','" +                                //TestName
                                    readings.ToString() + "','" +                                           //Reading
                                    (interval * 1000).ToString() + "','" +                                  //interval in msec
                                    slaveRow.ToString() + "','" +                                           // slaveRow number
                                    d.Rows[slaveRow][10].ToString() + "',#" +                               // charger type
                                    DateTime.Now.Add(new TimeSpan(0, 0, 2)).ToString("yyyy-MM-dd HH:mm:ss") + "#,#" +                  // start date
                                    DateTime.Now.Add(new TimeSpan(0, 0, 2)).ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  // date completed
                                    comboText + "','" +                                                     // technician
                                    (GlobalVars.CScanData[slaveRow].terminalID + 216).ToString() + "','" +  //terminal ID
                                    GlobalVars.CScanData[slaveRow].CCID.ToString() + "','" +                //cells cable ID
                                    GlobalVars.CScanData[slaveRow].SHCID.ToString() + "','" +               //shunt cable ID
                                    GlobalVars.CScanData[slaveRow].TCAB.ToString() + "','" +                //temp cable ID
                                    d.Rows[slaveRow][9].ToString() + "','" +                                //charger ID (Terminal Number)
                                    GlobalVars.CScanData[slaveRow].batNumCable10.ToString() +               //BATNUMCABLE10
                                    "');";


                                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                                    lock (dataBaseLock)
                                    {
                                        myAccessConn.Open();
                                        myAccessCommand.ExecuteNonQuery();
                                        myAccessConn.Close();
                                    }
                                }

                            }
                            catch (Exception ex)
                            {
                                myAccessConn.Close();

                                updateD(station, 5, false);
                                if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                //clear values from d
                                updateD(station, 7, ("Error"));
                                if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                                clearTLock();
                                cRunTest[station].Cancel();
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                });
                                return;
                            }



                            //We made it to this point without errors, so we'll update the grid with the step number
                            updateD(slaveRow, 3, slaveStepNum.ToString());

                        }

                        #endregion
                    }// end if

                    else    // We've got a resume!
                    {
                        #region look up current test numbers

                        //get the current step number from d
                        stepNum = int.Parse((string)d.Rows[station][3]);

                        if (MWO2 != "")
                        {
                            try
                            {
                                strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + MWO2 + "' ORDER BY StepNumber DESC;";
                                DataSet tests = new DataSet();
                                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                                lock (dataBaseLock)
                                {
                                    myAccessConn.Open();
                                    myDataAdapter.Fill(tests, "Tests");
                                    myAccessConn.Close();
                                }

                                if (tests.Tables[0].Rows.Count == 0)
                                {
                                    stepNum2 = 1;
                                }
                                else
                                {
                                    stepNum2 = int.Parse((string)tests.Tables[0].Rows[0][4]); // set to last step
                                }

                                if (MWO3 != "")
                                {
                                    strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + MWO3 + "' ORDER BY StepNumber DESC;";
                                    tests = new DataSet();
                                    myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                                    myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                                    lock (dataBaseLock)
                                    {
                                        myAccessConn.Open();
                                        myDataAdapter.Fill(tests, "Tests");
                                        myAccessConn.Close();
                                    }

                                    if (tests.Tables[0].Rows.Count == 0)
                                    {
                                        stepNum3 = 1;
                                    }
                                    else
                                    {
                                        stepNum3 = int.Parse((string)tests.Tables[0].Rows[0][4]);  // set to last step
                                    }
                                }


                            }
                            catch (Exception ex)
                            {
                                myAccessConn.Close();

                                updateD(station, 5, false);
                                if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                //clear values from d
                                updateD(station, 7, ("Error"));
                                if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                                clearTLock();
                                cRunTest[station].Cancel();
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                });
                                return;
                            }
                        }


                        //do we need to do the same for the slave?
                        if (MasterSlaveTest)
                        {
                            slaveStepNum = int.Parse((string)d.Rows[slaveRow][3]);

                            if (SWO2 != "")
                            {
                                // Now we'll look up the current test number and increment the new step number/////////////////////////////////////////////////////
                                //  open the db and pull in the options table
                                try
                                {
                                    strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + SWO2 + "' ORDER BY StepNumber DESC;";
                                    DataSet tests = new DataSet();
                                    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                                    lock (dataBaseLock)
                                    {
                                        myAccessConn.Open();
                                        myDataAdapter.Fill(tests, "Tests");
                                        myAccessConn.Close();
                                    }

                                    if (tests.Tables[0].Rows.Count == 0)
                                    {
                                        slaveStepNum2 = 1;
                                    }
                                    else
                                    {
                                        slaveStepNum2 = int.Parse((string)tests.Tables[0].Rows[0][4]);  // set to last step
                                    }

                                    if (SWO3 != "")
                                    {
                                        strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + SWO3 + "' ORDER BY StepNumber DESC;";
                                        tests = new DataSet();
                                        myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                                        myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                                        lock (dataBaseLock)
                                        {
                                            myAccessConn.Open();
                                            myDataAdapter.Fill(tests, "Tests");
                                            myAccessConn.Close();
                                        }

                                        if (tests.Tables[0].Rows.Count == 0)
                                        {
                                            slaveStepNum3 = 1;
                                        }
                                        else
                                        {
                                            slaveStepNum3 = int.Parse((string)tests.Tables[0].Rows[0][4]);  // set to last step
                                        }
                                    }

                                }
                                catch (Exception ex)
                                {
                                    myAccessConn.Close();
                                    updateD(station, 5, false);
                                    if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                    //clear values from d
                                    updateD(station, 7, ("Error"));
                                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                                    clearTLock();
                                    cRunTest[station].Cancel();
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                                        MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    });
                                    return;
                                }
                            }
                        }
                        #endregion
                    }
                    #endregion


                    // We should be ready to go at this point!
                    // and indicate that the test is starting


                    //reset the menu...
                    this.Invoke((MethodInvoker)delegate()
                    {
                        startNewTestToolStripMenuItem.Enabled = false;
                        resumeTestToolStripMenuItem.Enabled = false;
                        stopTestToolStripMenuItem.Enabled = true;
                    });

                    #region timer setup
                    // We are now good to go on starting the test loop timer...
                    // going to do the timming with a stop watch
                    bool firstRun = true;  // so we know if we should call fillPlotCombos()
                    int currentReading;
                    string oldETime = "";
                    var stopwatch = new Stopwatch();
                    TimeSpan offset;

                    //first check if we are resuming
                    if ((string)d.Rows[station][6] != "")
                    {
                        updateD(station, 7, "Resuming Test");
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Resuming Test"); }
                        this.Invoke((MethodInvoker)delegate() { sendNote(station, 3, "Test Resumed"); });
                        // we got a resume!
                        string temp = (string)d.Rows[station][6];
                        offset = new TimeSpan(int.Parse(temp.Substring(0, 2)), int.Parse(temp.Substring(3, 2)), int.Parse(temp.Substring(6, 2)));
                        currentReading = ((offset.Hours * 3600 + offset.Minutes * 60 + offset.Seconds) / interval) + 2;
                    }
                    else
                    {
                        updateD(station, 7, "Starting Test");
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Starting Test"); }
                        this.Invoke((MethodInvoker)delegate() { sendNote(station, 3, "Test Initiated -> " + testType); });
                        //fresh test!
                        offset = new TimeSpan();
                        currentReading = 1;
                    }

                    TimeSpan eTime = new TimeSpan().Add(offset);
                    string eTimeS = eTime.ToString(@"hh\:mm\:ss");

                    Thread.Sleep(2000);  // here so that we can actually see the grid update


                    #endregion

                    //cancel test
                    #region cancel block
                    if (token.IsCancellationRequested)
                    {

                        if (testType == "As Received" || testType.Contains("Shorting"))
                        {
                            //nothing to do here...
                        }
                        //clear values from d
                        updateD(station, 7, "Test Cancelled");
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Test Cancelled"); }
                        updateD(station, 5, false);
                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                        //update the gui
                        this.Invoke((MethodInvoker)delegate()
                        {
                            sendNote(station, 3, "Test Cancelled");
                            startNewTestToolStripMenuItem.Enabled = true;
                            resumeTestToolStripMenuItem.Enabled = true;
                            stopTestToolStripMenuItem.Enabled = false;
                        });

                        clearTLock();
                        cRunTest[station].Cancel();
                        updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                        return;
                    }
                    #endregion

                    #endregion

                    /////////////////////////////////START UP//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    #region START UP
                    // here for sinking up the ICs
                    bool firstCheat = false;

                    if (testType == "As Received" || testType.Contains("Shorting"))
                    {
                        // nothing to do! if it's an "As Received" or a shorting test... 
                        stopwatch.Start();
                    }
                    else if ((d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) && !runAsShunt)
                    {
                        #region             mode test
                        //update the GUI and pause...
                        updateD(station, 7, "Confirming Mode");
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Confirming Mode"); }
                        Thread.Sleep(2000);

                        // we need to check that we are in the correct mode for the test...
                        string temp = GlobalVars.ICData[Cstation].testMode.ToString();
                        switch (testType)
                        {
                            case "Full Charge-6":
                            case "Full Charge-4":
                                if (!(temp.Contains("20") || temp.Contains("21")))
                                {
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        DialogResult dialogResult = MessageBox.Show(this, "The charger doesn't seem to be set up for a Full Charge. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                        if (dialogResult == DialogResult.No)
                                        {
                                            cRunTest[station].Cancel();
                                            return;
                                        }
                                    });
                                    if (cRunTest[station].IsCancellationRequested == true)
                                    {

                                        updateD(station, 5, false);
                                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                        updateD(station, 7, "");
                                        if (MasterSlaveTest) { updateD(slaveRow, 7, ""); }
                                        //update the gui
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            startNewTestToolStripMenuItem.Enabled = true;
                                            resumeTestToolStripMenuItem.Enabled = true;
                                            stopTestToolStripMenuItem.Enabled = false;
                                        });
                                        clearTLock();
                                        cRunTest[station].Cancel();
                                        updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                        return;
                                    }  // end if

                                } // end if
                                break;
                            case "Top Charge-4":
                            case "Top Charge-2":
                            case "Top Charge-1":
                                if (!(temp.Contains("10") || temp.Contains("11")))
                                {
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        DialogResult dialogResult = MessageBox.Show(this, "The charger doesn't seem to be set up for Top Charge. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                        if (dialogResult == DialogResult.No)
                                        {
                                            cRunTest[station].Cancel();
                                            return;
                                        }
                                    });
                                    if (cRunTest[station].IsCancellationRequested == true)
                                    {
                                        updateD(station, 5, false);
                                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                        updateD(station, 7, "");
                                        if (MasterSlaveTest) { updateD(slaveRow, 7, ""); }
                                        //update the gui
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            startNewTestToolStripMenuItem.Enabled = true;
                                            resumeTestToolStripMenuItem.Enabled = true;
                                            stopTestToolStripMenuItem.Enabled = false;
                                        });
                                        clearTLock();
                                        cRunTest[station].Cancel();
                                        updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                        return;
                                    }  // end if
                                } // end if
                                break;
                            case "Constant Voltage":
                                if (!temp.Contains("12"))
                                {
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        DialogResult dialogResult = MessageBox.Show(this, "The charger doesn't seem to be set up for Constant Voltage. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                        if (dialogResult == DialogResult.No)
                                        {
                                            cRunTest[station].Cancel();
                                            return;
                                        }
                                    });
                                    if (cRunTest[station].IsCancellationRequested == true)
                                    {
                                        updateD(station, 5, false);
                                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                        updateD(station, 7, "");
                                        if (MasterSlaveTest) { updateD(slaveRow, 7, ""); }
                                        //update the gui
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            startNewTestToolStripMenuItem.Enabled = true;
                                            resumeTestToolStripMenuItem.Enabled = true;
                                            stopTestToolStripMenuItem.Enabled = false;
                                        });
                                        clearTLock();
                                        cRunTest[station].Cancel();
                                        updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                        return;
                                    }  // end if
                                } // end if
                                break;
                            case "Capacity-1":
                                if (!(temp.Contains("31") || temp.Contains("32")))
                                {
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        DialogResult dialogResult = MessageBox.Show(this, "The charger doesn't seem to be set up for a Capacity test. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                        if (dialogResult == DialogResult.No)
                                        {
                                            cRunTest[station].Cancel();
                                            return;
                                        }
                                    });
                                    if (cRunTest[station].IsCancellationRequested == true)
                                    {
                                        updateD(station, 5, false);
                                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                        updateD(station, 7, "");
                                        if (MasterSlaveTest) { updateD(slaveRow, 7, ""); }
                                        //update the gui
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            startNewTestToolStripMenuItem.Enabled = true;
                                            resumeTestToolStripMenuItem.Enabled = true;
                                            stopTestToolStripMenuItem.Enabled = false;
                                        });
                                        clearTLock();
                                        cRunTest[station].Cancel();
                                        updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                        return;
                                    }  // end if
                                } // end if
                                break;
                            case "Discharge":
                                if (!temp.Contains("30"))
                                {
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        DialogResult dialogResult = MessageBox.Show(this, "The charger doesn't seem to be set up for Discharge. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                        if (dialogResult == DialogResult.No)
                                        {
                                            cRunTest[station].Cancel();
                                            return;
                                        }
                                    });
                                    if (cRunTest[station].IsCancellationRequested == true)
                                    {
                                        updateD(station, 5, false);
                                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                        updateD(station, 7, "");
                                        if (MasterSlaveTest) { updateD(slaveRow, 7, ""); }
                                        //update the gui
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            startNewTestToolStripMenuItem.Enabled = true;
                                            resumeTestToolStripMenuItem.Enabled = true;
                                            stopTestToolStripMenuItem.Enabled = false;
                                        });
                                        clearTLock();
                                        cRunTest[station].Cancel();
                                        updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                        return;
                                    }  // end if
                                } // end if
                                break;
                            case "Slow Charge-14":
                            case "Slow Charge-16":
                                if (!(temp.Contains("10") || temp.Contains("11")))
                                {
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        DialogResult dialogResult = MessageBox.Show(this, "The charger doesn't seem to be set up for Slow Charge. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                        if (dialogResult == DialogResult.No)
                                        {
                                            cRunTest[station].Cancel();
                                            return;
                                        }
                                    });
                                    if (cRunTest[station].IsCancellationRequested == true)
                                    {
                                        updateD(station, 5, false);
                                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                        updateD(station, 7, "");
                                        if (MasterSlaveTest) { updateD(slaveRow, 7, ""); }
                                        //update the gui
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            startNewTestToolStripMenuItem.Enabled = true;
                                            resumeTestToolStripMenuItem.Enabled = true;
                                            stopTestToolStripMenuItem.Enabled = false;
                                        });
                                        clearTLock();
                                        cRunTest[station].Cancel();
                                        updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                        return;
                                    }  // end if
                                } // end if
                                break;
                            case "Custom Chg":
                            case "Custom Chg 2":
                            case "Custom Chg 3":
                                if (!(temp.Contains("10") || temp.Contains("11") || temp.Contains("12") || temp.Contains("20") || temp.Contains("21")))
                                {
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        DialogResult dialogResult = MessageBox.Show(this, "The charger doesn't seem to be set up for a Custom Charge. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                        if (dialogResult == DialogResult.No)
                                        {
                                            cRunTest[station].Cancel();
                                            return;
                                        }
                                    });
                                    if (cRunTest[station].IsCancellationRequested == true)
                                    {
                                        updateD(station, 5, false);
                                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                        updateD(station, 7, "");
                                        if (MasterSlaveTest) { updateD(slaveRow, 7, ""); }
                                        //update the gui
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            startNewTestToolStripMenuItem.Enabled = true;
                                            resumeTestToolStripMenuItem.Enabled = true;
                                            stopTestToolStripMenuItem.Enabled = false;
                                        });
                                        clearTLock();
                                        cRunTest[station].Cancel();
                                        updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                        return;
                                    }  // end if
                                } // end if
                                break;
                            case "Custom Cap":
                                if (!(temp.Contains("30") || temp.Contains("31") || temp.Contains("32")))
                                {
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        DialogResult dialogResult = MessageBox.Show(this, "The charger doesn't seem to be set up for a Custom Capacity test. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                        if (dialogResult == DialogResult.No)
                                        {
                                            cRunTest[station].Cancel();
                                            return;
                                        }
                                    });
                                    if (cRunTest[station].IsCancellationRequested == true)
                                    {
                                        updateD(station, 5, false);
                                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                        updateD(station, 7, "");
                                        if (MasterSlaveTest) { updateD(slaveRow, 7, ""); }
                                        //update the gui
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            startNewTestToolStripMenuItem.Enabled = true;
                                            resumeTestToolStripMenuItem.Enabled = true;
                                            stopTestToolStripMenuItem.Enabled = false;
                                            sendNote(station, 1, "Test Mode Incorrect");
                                        });
                                        clearTLock();
                                        cRunTest[station].Cancel();
                                        updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                        return;
                                    }  // end if
                                } // end if
                                break;
                            default:
                                break;
                        }
                        #endregion


                        //cancel test
                        #region cancel block
                        if (token.IsCancellationRequested)
                        {
                            //clear values from d
                            updateD(station, 7, "Test Cancelled");
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Test Cancelled"); }
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                            //update the gui
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 2, "Test Cancelled");
                                startNewTestToolStripMenuItem.Enabled = true;
                                resumeTestToolStripMenuItem.Enabled = true;
                                stopTestToolStripMenuItem.Enabled = false;
                            });

                            clearTLock();
                            cRunTest[station].Cancel();
                            updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                            return;
                        }
                        #endregion

                        // If we are in hold and we are starting a new test we need to reset before starting!
                        if ((string)d.Rows[station][11] != "RESET" && (string)d.Rows[station][6] == "")
                        {
                            for (int j = 0; j < 10; j++)
                            {
                                // now we need to reset the charger
                                updateD(station, 7, "Resetting Charger");
                                if (MasterSlaveTest) { updateD(slaveRow, 7, "Resetting Charger"); }
                                // set KE1 to 2 ("command")
                                GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                                // set KE3 to RESET
                                GlobalVars.ICSettings[Cstation].KE3 = (byte)3;
                                //Update the output string value
                                GlobalVars.ICSettings[Cstation].UpdateOutText();
                                criticalNum[Cstation] = true;
                                //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                                Thread.Sleep(5000);
                                // set KE1 to 1 ("query")
                                GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                                //Update the output string value
                                GlobalVars.ICSettings[Cstation].UpdateOutText();

                                //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                                for (int i = 0; i < 15; i++)
                                {
                                    criticalNum[Cstation] = true;
                                    Thread.Sleep(1000);
                                    if (d.Rows[station][11].ToString() == "RESET")
                                    {
                                        break;
                                    }
                                }
                                if (d.Rows[station][11].ToString() == "RESET")
                                {
                                    break;
                                }
                            }

                        }

                        //cancel test
                        #region cancel block
                        if (token.IsCancellationRequested)
                        {
                            //clear values from d
                            updateD(station, 7, "Test Cancelled");
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Test Cancelled"); }
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                            //update the gui
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 2, "Test Cancelled");
                                startNewTestToolStripMenuItem.Enabled = true;
                                resumeTestToolStripMenuItem.Enabled = true;
                                stopTestToolStripMenuItem.Enabled = false;
                            });
                            clearTLock();
                            cRunTest[station].Cancel();
                            updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                            return;
                        }
                        #endregion

                        // we are not in hold and we had a fault that we needed to clear...
                        else if ((string)d.Rows[station][11] != "HOLD" && (string)d.Rows[station][6] != "")
                        {
                            // we are resuming after a fault has been corrected..
                            // now we need to reset the charger
                            updateD(station, 7, "Clearing Charger");
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Clearing Charger"); }


                            for (int j = 0; j < 10; j++)
                            {
                                // set KE1 to 2 ("command")
                                GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                                // set KE3 to clear
                                GlobalVars.ICSettings[Cstation].KE3 = (byte)0;
                                //Update the output string value
                                GlobalVars.ICSettings[Cstation].UpdateOutText();
                                criticalNum[Cstation] = true;
                                Thread.Sleep(3000);
                                // set KE1 to 1 ("query")
                                GlobalVars.ICSettings[Cstation].KE1 = (byte)0;

                                //Update the output string value
                                GlobalVars.ICSettings[Cstation].UpdateOutText();
                                for (int i = 0; i < 3; i++)
                                {
                                    criticalNum[Cstation] = true;
                                    Thread.Sleep(1000);
                                    if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "RESET")
                                    {
                                        break;
                                    }
                                }

                                if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "RESET")
                                {
                                    break;
                                }
                            }

                        }

                        //cancel test
                        #region cancel block
                        if (token.IsCancellationRequested)
                        {
                            //clear values from d
                            updateD(station, 7, "Test Cancelled");
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Test Cancelled"); }
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                            //update the gui
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 2, "Test Cancelled");
                                startNewTestToolStripMenuItem.Enabled = true;
                                resumeTestToolStripMenuItem.Enabled = true;
                                stopTestToolStripMenuItem.Enabled = false;
                            });

                            clearTLock();
                            cRunTest[station].Cancel();
                            updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                            return;
                        }
                        #endregion

                        updateD(station, 7, "Telling Charger to Run");
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Telling Charger to Run"); }
                        Thread.Sleep(2000);
                        // we are going to use a thread in a thread to start the charger up
                        // maybe this will reduce the timing difference!

                        // set KE1 to 2 ("command")
                        GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                        // set KE3 to run
                        GlobalVars.ICSettings[Cstation].KE3 = (byte)1;
                        //Update the output string value
                        GlobalVars.ICSettings[Cstation].UpdateOutText();

                        for (int j = 0; j < 10; j++)
                        {
                            criticalNum[Cstation] = true;
                            Thread.Sleep(100);

                            for (int i = 0; i < 30; i++)
                            {
                                criticalNum[Cstation] = true;
                                Thread.Sleep(100);
                                if (d.Rows[station][11].ToString() != "HOLD" && d.Rows[station][11].ToString() != "RESET")
                                {
                                    break;
                                }
                            }

                            //make sure the charger has priority
                            if (d.Rows[station][11].ToString() != "HOLD" && d.Rows[station][11].ToString() != "RESET")
                            {
                                break;
                            }
                            //cancel test
                            #region cancel block
                            if (token.IsCancellationRequested)
                            {
                                updateD(station, 7, "Telling Charger to Stop");
                                if (MasterSlaveTest) { updateD(slaveRow, 7, "Telling Charger to Stop"); }

                                for (int jj = 0; jj < 3; jj++)
                                {

                                    // set KE1 to 2 ("command")
                                    GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                                    // set KE3 to stop
                                    GlobalVars.ICSettings[Cstation].KE3 = (byte)2;
                                    //Update the output string value
                                    GlobalVars.ICSettings[Cstation].UpdateOutText();
                                    criticalNum[Cstation] = true;
                                    //now we are going to create a thread to set KE1 back to data mode after 15 seconds

                                    for (int ii = 0; ii < 7; ii++)
                                    {
                                        criticalNum[Cstation] = true;
                                        Thread.Sleep(1000);
                                        if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "RESET")
                                        {
                                            break;
                                        }

                                    }
                                    if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "RESET")
                                    {
                                        // set KE1 to 0 ("query")
                                        GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                                        //Update the output string value
                                        GlobalVars.ICSettings[Cstation].UpdateOutText();
                                        break;
                                    }
                                }

                                //clear values from d
                                updateD(station, 7, "Test Cancelled");
                                if (MasterSlaveTest) { updateD(slaveRow, 7, "Test Cancelled"); }
                                updateD(station, 5, false);
                                if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                                //update the gui
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    sendNote(station, 2, "Test Cancelled");
                                    startNewTestToolStripMenuItem.Enabled = true;
                                    resumeTestToolStripMenuItem.Enabled = true;
                                    stopTestToolStripMenuItem.Enabled = false;
                                });

                                clearTLock();
                                cRunTest[station].Cancel();
                                updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                return;
                            }
                            #endregion
                        }

                        // set KE1 to 1 ("query")
                        GlobalVars.ICSettings[Cstation].KE1 = (byte)0;

                        //Update the output string value
                        GlobalVars.ICSettings[Cstation].UpdateOutText();

                        // add 2.5 seconds to help out the ICAs (don't need for MFCs)
                        if (!d.Rows[station][10].ToString().Contains("MFC"))
                        {
                            firstCheat = true;
                            offset = offset.Add(new TimeSpan(0, 0, 0, 0, 2500));
                        }
                        stopwatch.Start();

                        //Thread.Sleep(500);


                    }// end else if for ICs...
                    else if (d.Rows[station][10].ToString().Contains("CCA") && !runAsShunt)
                    {
                        // We have a legacy Charger!
                        // We need to let it run!

                        //cancel test
                        #region cancel block
                        if (token.IsCancellationRequested)
                        {
                            //clear values from d
                            updateD(station, 7, "Test Cancelled");
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Test Cancelled"); }
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                            cRunTest[station].Cancel();
                            //update the gui
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 2, "Test Cancelled");
                                startNewTestToolStripMenuItem.Enabled = true;
                                resumeTestToolStripMenuItem.Enabled = true;
                                stopTestToolStripMenuItem.Enabled = false;
                            });

                            //update the finish time
                            updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                            return;
                        }
                        #endregion

                        stopwatch.Start();
                        Thread.Sleep(500);
                        GlobalVars.cHold[station] = false;
                    }  // end else if for Legacy chargers
                    else if (d.Rows[station][10].ToString().Contains("Shunt") || runAsShunt)
                    {
                        // We have a shunt!!!!!!
                        // We'll start the test when we start to see current...
                        updateD(station, 7, "Waiting to see a current!");
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Waiting to see a current!"); }
                        while (true)
                        {
                            // make sure we didn't get a cancel first...
                            #region cancel block
                            if (token.IsCancellationRequested)
                            {
                                //clear values from d
                                updateD(station, 7, ("Cancelled"));
                                if (MasterSlaveTest) { updateD(slaveRow, 7, "Cancelled"); }
                                updateD(station, 5, false);
                                if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                cRunTest[station].Cancel();
                                //update the gui
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    sendNote(station, 2, "Test Cancelled");
                                    startNewTestToolStripMenuItem.Enabled = true;
                                    resumeTestToolStripMenuItem.Enabled = true;
                                    stopTestToolStripMenuItem.Enabled = false;
                                });

                                clearTLock();
                                //update the finish time
                                updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                return;
                            }
                            #endregion
                            // look for a current
                            if (Math.Abs(GlobalVars.CScanData[station].currentOne) > 0.2 || GlobalVars.allowZeroShunt)
                            {
                                // we found a current!
                                stopwatch.Start();
                                break;
                            }
                            Thread.Sleep(100);
                        }
                    }  // end the shunt else!
                    else
                    {
                        // We don't have a charger linked ...
                        //clear values from d
                        updateD(station, 7, ("No Charger Detected!"));
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "No Charger Detected!"); }
                        updateD(station, 5, false);
                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                        //update the gui
                        this.Invoke((MethodInvoker)delegate()
                        {
                            sendNote(station, 1, "No Charger Detected");
                            startNewTestToolStripMenuItem.Enabled = true;
                            resumeTestToolStripMenuItem.Enabled = true;
                            stopTestToolStripMenuItem.Enabled = false;
                        });

                        clearTLock();   // for good luck!
                        cRunTest[station].Cancel();
                        //update the finish time
                        updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "No Charger Detected!  Please check settings!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        return;
                    }

                    // we made it.
                    //tell them what reading we are on!
                    updateD(station, 7, ("Reading " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                    if (MasterSlaveTest) { updateD(slaveRow, 7, ("Reading " + (currentReading - 1).ToString() + " of " + readings.ToString())); }

                    bool startUpDelay = true;
                    // clear startUpDelay in 20 seconds...
                    // the start up delay is here to inhibit the error checking during startup...
                    ThreadPool.QueueUserWorkItem(t =>
                    {
                        Thread.Sleep(20000);
                        startUpDelay = false;
                    });


                    clearTLock();

                    #endregion

                    /////////////////////////////////MAIN LOOP//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    #region MAIN LOOP

                    //used for counting the number of bad current readings...
                    int badCurCount = 0;
                    byte badReadingCount = 0;
                    //used for looking for temperature drift of temp plate
                    double startTemp = GlobalVars.CScanData[station].TP1 + GlobalVars.CScanData[station].TP2 + GlobalVars.CScanData[station].TP3 + GlobalVars.CScanData[station].TP4;
                    startTemp /= 4;
                    int memTemp = 0;

                    while (currentReading <= readings)
                    {
                        // main loop try
                        try
                        {
                            // check if we need to take a reading
                            if (((currentReading - 1) * interval * 1000) < stopwatch.Elapsed.Add(offset).TotalMilliseconds || (GlobalVars.StopOnEnd && (d.Rows[station][11].ToString() == "END" || d.Rows[station][11].ToString() == "Peak End")))
                            {
                                //Compensate for the "END"
                                if (GlobalVars.StopOnEnd && (d.Rows[station][11].ToString() == "END" || d.Rows[station][11].ToString() == "Peak End"))
                                {
                                    currentReading = readings;
                                }

                                //first record the elapsed amount of time
                                TimeSpan temp = stopwatch.Elapsed.Add(offset);
                                if (firstCheat)
                                {
                                    firstCheat = false;
                                    temp = new TimeSpan(0, 0, 0);
                                }
                                // update the grid
                                updateD(station, 7, ("Reading " + currentReading.ToString() + " of " + readings.ToString()));
                                if (MasterSlaveTest) { updateD(slaveRow, 7, ("Reading " + currentReading.ToString() + " of " + readings.ToString())); }

                                #region save a scan to the DB and update the main screen plots
                                //  now try to INSERT INTO it
                                try
                                {
                                    string strUpdateCMD = "INSERT INTO ScanData (Station,BWO,STEP,RDG,[DATE],QS1,CTR,ETIME,CUR1,CUR2,VB1,VB2,VB3,VB4," +
                                        "CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24," +
                                        "BT1,BT2,BT3,BT4,BT5,BT6,CGND1,CGND2,REF,GND,FV,MSV,PSV) "
                                        + "VALUES (" + station.ToString() + ",'" +                            //station number
                                        MWO1.Trim() + "','" +                          //WorkOrderNumber
                                        stepNum.ToString("00") + "'," +                                            //StepNumber
                                        currentReading.ToString() + ",#" +                                     //ReadingNumber
                                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  //date
                                        GlobalVars.CScanData[station].QS1.ToString() + "','" +                  //QS1
                                        GlobalVars.CScanData[station].CTR.ToString() + "','" +                  //CTR
                                        temp.TotalDays.ToString("0.00000") + "','" +                                      //time elapsed in days
                                        GlobalVars.CScanData[station].currentOne.ToString("0.0") + "','" +           //CUR1
                                        GlobalVars.CScanData[station].currentTwo.ToString("0.0") + "','" +           //CUR2
                                        GlobalVars.CScanData[station].VB1.ToString("0.00") + "','" +                  //VB1
                                        GlobalVars.CScanData[station].VB2.ToString("0.00") + "','" +                  //VB2
                                        GlobalVars.CScanData[station].VB3.ToString("0.00") + "','" +                  //VB3
                                        GlobalVars.CScanData[station].VB4.ToString("0.00") + "','" +                  //VB4
                                        GlobalVars.CScanData[station].orderedCells[0].ToString("0.000") + "','" +      //CEL01
                                        GlobalVars.CScanData[station].orderedCells[1].ToString("0.000") + "','" +      //CEL02
                                        GlobalVars.CScanData[station].orderedCells[2].ToString("0.000") + "','" +      //CEL03
                                        GlobalVars.CScanData[station].orderedCells[3].ToString("0.000") + "','" +      //CEL04
                                        GlobalVars.CScanData[station].orderedCells[4].ToString("0.000") + "','" +      //CEL05
                                        GlobalVars.CScanData[station].orderedCells[5].ToString("0.000") + "','" +      //CEL06
                                        GlobalVars.CScanData[station].orderedCells[6].ToString("0.000") + "','" +      //CEL07
                                        GlobalVars.CScanData[station].orderedCells[7].ToString("0.000") + "','" +      //CEL08
                                        GlobalVars.CScanData[station].orderedCells[8].ToString("0.000") + "','" +      //CEL09
                                        GlobalVars.CScanData[station].orderedCells[9].ToString("0.000") + "','" +      //CEL10
                                        GlobalVars.CScanData[station].orderedCells[10].ToString("0.000") + "','" +     //CEL11
                                        GlobalVars.CScanData[station].orderedCells[11].ToString("0.000") + "','" +     //CEL12
                                        GlobalVars.CScanData[station].orderedCells[12].ToString("0.000") + "','" +     //CEL13
                                        GlobalVars.CScanData[station].orderedCells[13].ToString("0.000") + "','" +     //CEL14
                                        GlobalVars.CScanData[station].orderedCells[14].ToString("0.000") + "','" +     //CEL15
                                        GlobalVars.CScanData[station].orderedCells[15].ToString("0.000") + "','" +     //CEL16
                                        GlobalVars.CScanData[station].orderedCells[16].ToString("0.000") + "','" +     //CEL17
                                        GlobalVars.CScanData[station].orderedCells[17].ToString("0.000") + "','" +     //CEL18
                                        GlobalVars.CScanData[station].orderedCells[18].ToString("0.000") + "','" +     //CEL19
                                        GlobalVars.CScanData[station].orderedCells[19].ToString("0.000") + "','" +     //CEL20
                                        GlobalVars.CScanData[station].orderedCells[20].ToString("0.000") + "','" +     //CEL21
                                        GlobalVars.CScanData[station].orderedCells[21].ToString("0.000") + "','" +     //CEL22
                                        GlobalVars.CScanData[station].orderedCells[22].ToString("0.000") + "','" +     //CEL23
                                        GlobalVars.CScanData[station].orderedCells[23].ToString("0.000") + "','" +     //CEL24
                                        GlobalVars.CScanData[station].TP1.ToString("0.0") + "','" +                  //TP1
                                        GlobalVars.CScanData[station].TP2.ToString("0.0") + "','" +                  //TP2
                                        GlobalVars.CScanData[station].TP3.ToString("0.0") + "','" +                  //TP3
                                        GlobalVars.CScanData[station].TP4.ToString("0.0") + "','" +                  //TP4
                                        GlobalVars.CScanData[station].TP5.ToString("0.0") + "','" +                  //TP5
                                        "0.0" + "','" +                                                         //TP6
                                        GlobalVars.CScanData[station].cellGND1.ToString("0.000") + "','" +             //CGND1
                                        GlobalVars.CScanData[station].cellGND2.ToString("0.000") + "','" +             //CGND2
                                        GlobalVars.CScanData[station].ref95V.ToString("0.000") + "','" +               //ref
                                        GlobalVars.CScanData[station].ch0GND.ToString("0.000") + "','" +               //GND
                                        GlobalVars.CScanData[station].plus5V.ToString("0.000") + "','" +              //FV
                                        GlobalVars.CScanData[station].minus15.ToString("0.00") + "','" +              //MSV
                                        GlobalVars.CScanData[station].plus15.ToString("0.00") +                       //PSV
                                        "');";

                                    OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                                    lock (dataBaseLock)
                                    {
                                        myAccessConn.Open();
                                        myAccessCommand.ExecuteNonQuery();
                                        myAccessConn.Close();
                                    }

                                    if (MWO2 != "" && MWO3 == "")
                                    {
                                        // this is the case where the second work order is the bottom 11 cells of a 2X11 cable...
                                        strUpdateCMD = "INSERT INTO ScanData (Station,BWO,STEP,RDG,[DATE],QS1,CTR,ETIME,CUR1,CUR2,VB1,VB2,VB3,VB4," +
                                        "CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24," +
                                        "BT1,BT2,BT3,BT4,BT5,BT6,CGND1,CGND2,REF,GND,FV,MSV,PSV) "
                                        + "VALUES (" + station.ToString() + ",'" +                            //station number
                                        MWO2.Trim() + "','" +                          //WorkOrderNumber
                                        stepNum2.ToString("00") + "'," +                                            //StepNumber
                                        currentReading.ToString() + ",#" +                                     //ReadingNumber
                                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  //date
                                        GlobalVars.CScanData[station].QS1.ToString() + "','" +                  //QS1
                                        GlobalVars.CScanData[station].CTR.ToString() + "','" +                  //CTR
                                        temp.TotalDays.ToString("0.00000") + "','" +                                      //time elapsed in days
                                        GlobalVars.CScanData[station].currentOne.ToString("0.0") + "','" +           //CUR1
                                        GlobalVars.CScanData[station].currentTwo.ToString("0.0") + "','" +           //CUR2
                                        GlobalVars.CScanData[station].VB1.ToString("0.00") + "','" +                  //VB1
                                        GlobalVars.CScanData[station].VB2.ToString("0.00") + "','" +                  //VB2
                                        GlobalVars.CScanData[station].VB3.ToString("0.00") + "','" +                  //VB3
                                        GlobalVars.CScanData[station].VB4.ToString("0.00") + "','" +                  //VB4
                                        GlobalVars.CScanData[station].orderedCells[12].ToString("0.000") + "','" +      //CEL01
                                        GlobalVars.CScanData[station].orderedCells[13].ToString("0.000") + "','" +      //CEL02
                                        GlobalVars.CScanData[station].orderedCells[14].ToString("0.000") + "','" +      //CEL03
                                        GlobalVars.CScanData[station].orderedCells[15].ToString("0.000") + "','" +      //CEL04
                                        GlobalVars.CScanData[station].orderedCells[16].ToString("0.000") + "','" +      //CEL05
                                        GlobalVars.CScanData[station].orderedCells[17].ToString("0.000") + "','" +      //CEL06
                                        GlobalVars.CScanData[station].orderedCells[18].ToString("0.000") + "','" +      //CEL07
                                        GlobalVars.CScanData[station].orderedCells[19].ToString("0.000") + "','" +      //CEL08
                                        GlobalVars.CScanData[station].orderedCells[20].ToString("0.000") + "','" +      //CEL09
                                        GlobalVars.CScanData[station].orderedCells[21].ToString("0.000") + "','" +      //CEL10
                                        GlobalVars.CScanData[station].orderedCells[22].ToString("0.000") + "','" +     //CEL11
                                        GlobalVars.CScanData[station].orderedCells[0].ToString("0.000") + "','" +     //CEL12
                                        GlobalVars.CScanData[station].orderedCells[1].ToString("0.000") + "','" +     //CEL13
                                        GlobalVars.CScanData[station].orderedCells[2].ToString("0.000") + "','" +     //CEL14
                                        GlobalVars.CScanData[station].orderedCells[3].ToString("0.000") + "','" +     //CEL15
                                        GlobalVars.CScanData[station].orderedCells[4].ToString("0.000") + "','" +     //CEL16
                                        GlobalVars.CScanData[station].orderedCells[5].ToString("0.000") + "','" +     //CEL17
                                        GlobalVars.CScanData[station].orderedCells[6].ToString("0.000") + "','" +     //CEL18
                                        GlobalVars.CScanData[station].orderedCells[7].ToString("0.000") + "','" +     //CEL19
                                        GlobalVars.CScanData[station].orderedCells[8].ToString("0.000") + "','" +     //CEL20
                                        GlobalVars.CScanData[station].orderedCells[9].ToString("0.000") + "','" +     //CEL21
                                        GlobalVars.CScanData[station].orderedCells[10].ToString("0.000") + "','" +     //CEL22
                                        GlobalVars.CScanData[station].orderedCells[11].ToString("0.000") + "','" +     //CEL23
                                        GlobalVars.CScanData[station].orderedCells[23].ToString("0.000") + "','" +     //CEL24
                                        GlobalVars.CScanData[station].TP1.ToString("0.0") + "','" +                  //TP1
                                        GlobalVars.CScanData[station].TP2.ToString("0.0") + "','" +                  //TP2
                                        GlobalVars.CScanData[station].TP3.ToString("0.0") + "','" +                  //TP3
                                        GlobalVars.CScanData[station].TP4.ToString("0.0") + "','" +                  //TP4
                                        GlobalVars.CScanData[station].TP5.ToString("0.0") + "','" +                  //TP5
                                        "0.0" + "','" +                                                         //TP6
                                        GlobalVars.CScanData[station].cellGND1.ToString("0.000") + "','" +             //CGND1
                                        GlobalVars.CScanData[station].cellGND2.ToString("0.000") + "','" +             //CGND2
                                        GlobalVars.CScanData[station].ref95V.ToString("0.000") + "','" +               //ref
                                        GlobalVars.CScanData[station].ch0GND.ToString("0.000") + "','" +               //GND
                                        GlobalVars.CScanData[station].plus5V.ToString("0.000") + "','" +              //FV
                                        GlobalVars.CScanData[station].minus15.ToString("0.00") + "','" +              //MSV
                                        GlobalVars.CScanData[station].plus15.ToString("0.00") +                       //PSV
                                        "');";

                                        myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                                        lock (dataBaseLock)
                                        {
                                            myAccessConn.Open();
                                            myAccessCommand.ExecuteNonQuery();
                                            myAccessConn.Close();
                                        }
                                    }
                                    else if (MWO2 != "")
                                    {
                                        // this is the case where the second work order is the middle 7 cells of a 3X7 cable...
                                        strUpdateCMD = "INSERT INTO ScanData (Station,BWO,STEP,RDG,[DATE],QS1,CTR,ETIME,CUR1,CUR2,VB1,VB2,VB3,VB4," +
                                        "CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24," +
                                        "BT1,BT2,BT3,BT4,BT5,BT6,CGND1,CGND2,REF,GND,FV,MSV,PSV) "
                                        + "VALUES (" + station.ToString() + ",'" +                            //station number
                                        MWO2.Trim() + "','" +                          //WorkOrderNumber
                                        stepNum2.ToString("00") + "'," +                                            //StepNumber
                                        currentReading.ToString() + ",#" +                                     //ReadingNumber
                                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  //date
                                        GlobalVars.CScanData[station].QS1.ToString() + "','" +                  //QS1
                                        GlobalVars.CScanData[station].CTR.ToString() + "','" +                  //CTR
                                        temp.TotalDays.ToString("0.00000") + "','" +                                      //time elapsed in days
                                        GlobalVars.CScanData[station].currentOne.ToString("0.0") + "','" +           //CUR1
                                        GlobalVars.CScanData[station].currentTwo.ToString("0.0") + "','" +           //CUR2
                                        GlobalVars.CScanData[station].VB1.ToString("0.00") + "','" +                  //VB1
                                        GlobalVars.CScanData[station].VB2.ToString("0.00") + "','" +                  //VB2
                                        GlobalVars.CScanData[station].VB3.ToString("0.00") + "','" +                  //VB3
                                        GlobalVars.CScanData[station].VB4.ToString("0.00") + "','" +                  //VB4
                                        GlobalVars.CScanData[station].orderedCells[8].ToString("0.000") + "','" +      //CEL01
                                        GlobalVars.CScanData[station].orderedCells[9].ToString("0.000") + "','" +      //CEL02
                                        GlobalVars.CScanData[station].orderedCells[10].ToString("0.000") + "','" +      //CEL03
                                        GlobalVars.CScanData[station].orderedCells[11].ToString("0.000") + "','" +      //CEL04
                                        GlobalVars.CScanData[station].orderedCells[12].ToString("0.000") + "','" +      //CEL05
                                        GlobalVars.CScanData[station].orderedCells[13].ToString("0.000") + "','" +      //CEL06
                                        GlobalVars.CScanData[station].orderedCells[14].ToString("0.000") + "','" +      //CEL07
                                        GlobalVars.CScanData[station].orderedCells[15].ToString("0.000") + "','" +      //CEL08
                                        GlobalVars.CScanData[station].orderedCells[16].ToString("0.000") + "','" +      //CEL09
                                        GlobalVars.CScanData[station].orderedCells[17].ToString("0.000") + "','" +      //CEL10
                                        GlobalVars.CScanData[station].orderedCells[18].ToString("0.000") + "','" +     //CEL11
                                        GlobalVars.CScanData[station].orderedCells[19].ToString("0.000") + "','" +     //CEL12
                                        GlobalVars.CScanData[station].orderedCells[20].ToString("0.000") + "','" +     //CEL13
                                        GlobalVars.CScanData[station].orderedCells[21].ToString("0.000") + "','" +     //CEL14
                                        GlobalVars.CScanData[station].orderedCells[22].ToString("0.000") + "','" +     //CEL15
                                        GlobalVars.CScanData[station].orderedCells[23].ToString("0.000") + "','" +     //CEL16
                                        GlobalVars.CScanData[station].orderedCells[0].ToString("0.000") + "','" +     //CEL17
                                        GlobalVars.CScanData[station].orderedCells[1].ToString("0.000") + "','" +     //CEL18
                                        GlobalVars.CScanData[station].orderedCells[2].ToString("0.000") + "','" +     //CEL19
                                        GlobalVars.CScanData[station].orderedCells[3].ToString("0.000") + "','" +     //CEL20
                                        GlobalVars.CScanData[station].orderedCells[4].ToString("0.000") + "','" +     //CEL21
                                        GlobalVars.CScanData[station].orderedCells[5].ToString("0.000") + "','" +     //CEL22
                                        GlobalVars.CScanData[station].orderedCells[6].ToString("0.000") + "','" +     //CEL23
                                        GlobalVars.CScanData[station].orderedCells[7].ToString("0.000") + "','" +     //CEL24
                                        GlobalVars.CScanData[station].TP1.ToString("0.0") + "','" +                  //TP1
                                        GlobalVars.CScanData[station].TP2.ToString("0.0") + "','" +                  //TP2
                                        GlobalVars.CScanData[station].TP3.ToString("0.0") + "','" +                  //TP3
                                        GlobalVars.CScanData[station].TP4.ToString("0.0") + "','" +                  //TP4
                                        GlobalVars.CScanData[station].TP5.ToString("0.0") + "','" +                  //TP5
                                        "0.0" + "','" +                                                         //TP6
                                        GlobalVars.CScanData[station].cellGND1.ToString("0.000") + "','" +             //CGND1
                                        GlobalVars.CScanData[station].cellGND2.ToString("0.000") + "','" +             //CGND2
                                        GlobalVars.CScanData[station].ref95V.ToString("0.000") + "','" +               //ref
                                        GlobalVars.CScanData[station].ch0GND.ToString("0.000") + "','" +               //GND
                                        GlobalVars.CScanData[station].plus5V.ToString("0.000") + "','" +              //FV
                                        GlobalVars.CScanData[station].minus15.ToString("0.00") + "','" +              //MSV
                                        GlobalVars.CScanData[station].plus15.ToString("0.00") +                       //PSV
                                        "');";

                                        myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                                        lock (dataBaseLock)
                                        {
                                            myAccessConn.Open();
                                            myAccessCommand.ExecuteNonQuery();
                                            myAccessConn.Close();
                                        }
                                    }

                                    if (MWO3 != "")
                                    {
                                        // this is the case where the third work order is the last 7 cells of a 3X7 cable...
                                        strUpdateCMD = "INSERT INTO ScanData (Station,BWO,STEP,RDG,[DATE],QS1,CTR,ETIME,CUR1,CUR2,VB1,VB2,VB3,VB4," +
                                        "CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24," +
                                        "BT1,BT2,BT3,BT4,BT5,BT6,CGND1,CGND2,REF,GND,FV,MSV,PSV) "
                                        + "VALUES (" + station.ToString() + ",'" +                            //station number
                                        MWO3.Trim() + "','" +                          //WorkOrderNumber
                                        stepNum3.ToString("00") + "'," +                                            //StepNumber
                                        currentReading.ToString() + ",#" +                                     //ReadingNumber
                                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  //date
                                        GlobalVars.CScanData[station].QS1.ToString() + "','" +                  //QS1
                                        GlobalVars.CScanData[station].CTR.ToString() + "','" +                  //CTR
                                        temp.TotalDays.ToString("0.00000") + "','" +                                      //time elapsed in days
                                        GlobalVars.CScanData[station].currentOne.ToString("0.0") + "','" +           //CUR1
                                        GlobalVars.CScanData[station].currentTwo.ToString("0.0") + "','" +           //CUR2
                                        GlobalVars.CScanData[station].VB1.ToString("0.00") + "','" +                  //VB1
                                        GlobalVars.CScanData[station].VB2.ToString("0.00") + "','" +                  //VB2
                                        GlobalVars.CScanData[station].VB3.ToString("0.00") + "','" +                  //VB3
                                        GlobalVars.CScanData[station].VB4.ToString("0.00") + "','" +                  //VB4
                                        GlobalVars.CScanData[station].orderedCells[16].ToString("0.000") + "','" +      //CEL01
                                        GlobalVars.CScanData[station].orderedCells[17].ToString("0.000") + "','" +      //CEL02
                                        GlobalVars.CScanData[station].orderedCells[18].ToString("0.000") + "','" +      //CEL03
                                        GlobalVars.CScanData[station].orderedCells[19].ToString("0.000") + "','" +      //CEL04
                                        GlobalVars.CScanData[station].orderedCells[20].ToString("0.000") + "','" +      //CEL05
                                        GlobalVars.CScanData[station].orderedCells[21].ToString("0.000") + "','" +      //CEL06
                                        GlobalVars.CScanData[station].orderedCells[22].ToString("0.000") + "','" +      //CEL07
                                        GlobalVars.CScanData[station].orderedCells[23].ToString("0.000") + "','" +      //CEL08
                                        GlobalVars.CScanData[station].orderedCells[0].ToString("0.000") + "','" +      //CEL09
                                        GlobalVars.CScanData[station].orderedCells[1].ToString("0.000") + "','" +      //CEL10
                                        GlobalVars.CScanData[station].orderedCells[2].ToString("0.000") + "','" +     //CEL11
                                        GlobalVars.CScanData[station].orderedCells[3].ToString("0.000") + "','" +     //CEL12
                                        GlobalVars.CScanData[station].orderedCells[4].ToString("0.000") + "','" +     //CEL13
                                        GlobalVars.CScanData[station].orderedCells[5].ToString("0.000") + "','" +     //CEL14
                                        GlobalVars.CScanData[station].orderedCells[6].ToString("0.000") + "','" +     //CEL15
                                        GlobalVars.CScanData[station].orderedCells[7].ToString("0.000") + "','" +     //CEL16
                                        GlobalVars.CScanData[station].orderedCells[8].ToString("0.000") + "','" +     //CEL17
                                        GlobalVars.CScanData[station].orderedCells[9].ToString("0.000") + "','" +     //CEL18
                                        GlobalVars.CScanData[station].orderedCells[10].ToString("0.000") + "','" +     //CEL19
                                        GlobalVars.CScanData[station].orderedCells[11].ToString("0.000") + "','" +     //CEL20
                                        GlobalVars.CScanData[station].orderedCells[12].ToString("0.000") + "','" +     //CEL21
                                        GlobalVars.CScanData[station].orderedCells[13].ToString("0.000") + "','" +     //CEL22
                                        GlobalVars.CScanData[station].orderedCells[14].ToString("0.000") + "','" +     //CEL23
                                        GlobalVars.CScanData[station].orderedCells[15].ToString("0.000") + "','" +     //CEL24
                                        GlobalVars.CScanData[station].TP1.ToString("0.0") + "','" +                  //TP1
                                        GlobalVars.CScanData[station].TP2.ToString("0.0") + "','" +                  //TP2
                                        GlobalVars.CScanData[station].TP3.ToString("0.0") + "','" +                  //TP3
                                        GlobalVars.CScanData[station].TP4.ToString("0.0") + "','" +                  //TP4
                                        GlobalVars.CScanData[station].TP5.ToString("0.0") + "','" +                  //TP5
                                        "0.0" + "','" +                                                         //TP6
                                        GlobalVars.CScanData[station].cellGND1.ToString("0.000") + "','" +             //CGND1
                                        GlobalVars.CScanData[station].cellGND2.ToString("0.000") + "','" +             //CGND2
                                        GlobalVars.CScanData[station].ref95V.ToString("0.000") + "','" +               //ref
                                        GlobalVars.CScanData[station].ch0GND.ToString("0.000") + "','" +               //GND
                                        GlobalVars.CScanData[station].plus5V.ToString("0.000") + "','" +              //FV
                                        GlobalVars.CScanData[station].minus15.ToString("0.00") + "','" +              //MSV
                                        GlobalVars.CScanData[station].plus15.ToString("0.00") +                       //PSV
                                        "');";

                                        myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                                        lock (dataBaseLock)
                                        {
                                            myAccessConn.Open();
                                            myAccessCommand.ExecuteNonQuery();
                                            myAccessConn.Close();
                                        }
                                    }

                                    //also insert the slave reading if need be...
                                    if (MasterSlaveTest)
                                    {

                                        strUpdateCMD = "INSERT INTO ScanData (Station,BWO,STEP,RDG,[DATE],QS1,CTR,ETIME,CUR1,CUR2,VB1,VB2,VB3,VB4," +
                                            "CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24," +
                                            "BT1,BT2,BT3,BT4,BT5,BT6,CGND1,CGND2,REF,GND,FV,MSV,PSV) "
                                            + "VALUES (" + slaveRow.ToString() + ",'" +                            //slaveRow number
                                            SWO1.Trim() + "','" +                          //WorkOrderNumber
                                            slaveStepNum.ToString("00") + "'," +                                            //StepNumber
                                            currentReading.ToString() + ",#" +                                     //ReadingNumber
                                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  //date
                                            GlobalVars.CScanData[slaveRow].QS1.ToString() + "','" +                  //QS1
                                            GlobalVars.CScanData[slaveRow].CTR.ToString() + "','" +                  //CTR
                                            temp.TotalDays.ToString("0.00000") + "','" +                                      //time elapsed in days
                                            GlobalVars.CScanData[station].currentOne.ToString("0.0") + "','" +           //CUR1  (pulled from master CSCAN)
                                            GlobalVars.CScanData[station].currentTwo.ToString("0.0") + "','" +           //CUR2  
                                            GlobalVars.CScanData[slaveRow].VB1.ToString("0.00") + "','" +                  //VB1
                                            GlobalVars.CScanData[slaveRow].VB2.ToString("0.00") + "','" +                  //VB2
                                            GlobalVars.CScanData[slaveRow].VB3.ToString("0.00") + "','" +                  //VB3
                                            GlobalVars.CScanData[slaveRow].VB4.ToString("0.00") + "','" +                  //VB4
                                            GlobalVars.CScanData[slaveRow].orderedCells[0].ToString("0.000") + "','" +      //CEL01
                                            GlobalVars.CScanData[slaveRow].orderedCells[1].ToString("0.000") + "','" +      //CEL02
                                            GlobalVars.CScanData[slaveRow].orderedCells[2].ToString("0.000") + "','" +      //CEL03
                                            GlobalVars.CScanData[slaveRow].orderedCells[3].ToString("0.000") + "','" +      //CEL04
                                            GlobalVars.CScanData[slaveRow].orderedCells[4].ToString("0.000") + "','" +      //CEL05
                                            GlobalVars.CScanData[slaveRow].orderedCells[5].ToString("0.000") + "','" +      //CEL06
                                            GlobalVars.CScanData[slaveRow].orderedCells[6].ToString("0.000") + "','" +      //CEL07
                                            GlobalVars.CScanData[slaveRow].orderedCells[7].ToString("0.000") + "','" +      //CEL08
                                            GlobalVars.CScanData[slaveRow].orderedCells[8].ToString("0.000") + "','" +      //CEL09
                                            GlobalVars.CScanData[slaveRow].orderedCells[9].ToString("0.000") + "','" +      //CEL10
                                            GlobalVars.CScanData[slaveRow].orderedCells[10].ToString("0.000") + "','" +     //CEL11
                                            GlobalVars.CScanData[slaveRow].orderedCells[11].ToString("0.000") + "','" +     //CEL12
                                            GlobalVars.CScanData[slaveRow].orderedCells[12].ToString("0.000") + "','" +     //CEL13
                                            GlobalVars.CScanData[slaveRow].orderedCells[13].ToString("0.000") + "','" +     //CEL14
                                            GlobalVars.CScanData[slaveRow].orderedCells[14].ToString("0.000") + "','" +     //CEL15
                                            GlobalVars.CScanData[slaveRow].orderedCells[15].ToString("0.000") + "','" +     //CEL16
                                            GlobalVars.CScanData[slaveRow].orderedCells[16].ToString("0.000") + "','" +     //CEL17
                                            GlobalVars.CScanData[slaveRow].orderedCells[17].ToString("0.000") + "','" +     //CEL18
                                            GlobalVars.CScanData[slaveRow].orderedCells[18].ToString("0.000") + "','" +     //CEL19
                                            GlobalVars.CScanData[slaveRow].orderedCells[19].ToString("0.000") + "','" +     //CEL20
                                            GlobalVars.CScanData[slaveRow].orderedCells[20].ToString("0.000") + "','" +     //CEL21
                                            GlobalVars.CScanData[slaveRow].orderedCells[21].ToString("0.000") + "','" +     //CEL22
                                            GlobalVars.CScanData[slaveRow].orderedCells[22].ToString("0.000") + "','" +     //CEL23
                                            GlobalVars.CScanData[slaveRow].orderedCells[23].ToString("0.000") + "','" +     //CEL24
                                            GlobalVars.CScanData[station].TP1.ToString("0.0") + "','" +                  //TP1  (pulled from master CSCAN)
                                            GlobalVars.CScanData[station].TP2.ToString("0.0") + "','" +                  //TP2  (pulled from master CSCAN)
                                            GlobalVars.CScanData[station].TP3.ToString("0.0") + "','" +                  //TP3  (pulled from master CSCAN)
                                            GlobalVars.CScanData[station].TP4.ToString("0.0") + "','" +                  //TP4  (pulled from master CSCAN)
                                            GlobalVars.CScanData[station].TP5.ToString("0.0") + "','" +                  //TP5  (pulled from master CSCAN)
                                            "0.0" + "','" +                                                         //TP6
                                            GlobalVars.CScanData[slaveRow].cellGND1.ToString("0.000") + "','" +             //CGND1
                                            GlobalVars.CScanData[slaveRow].cellGND2.ToString("0.000") + "','" +             //CGND2
                                            GlobalVars.CScanData[slaveRow].ref95V.ToString("0.000") + "','" +               //ref
                                            GlobalVars.CScanData[slaveRow].ch0GND.ToString("0.000") + "','" +               //GND
                                            GlobalVars.CScanData[slaveRow].plus5V.ToString("0.000") + "','" +              //FV
                                            GlobalVars.CScanData[slaveRow].minus15.ToString("0.00") + "','" +              //MSV
                                            GlobalVars.CScanData[slaveRow].plus15.ToString("0.00") +                       //PSV
                                            "');";

                                        myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                                        lock (dataBaseLock)
                                        {
                                            myAccessConn.Open();
                                            myAccessCommand.ExecuteNonQuery();
                                            myAccessConn.Close();
                                        } // end lock

                                        if (SWO2 != "" && SWO3 == "")
                                        {
                                            // this is the case where the second work order is the bottom 11 cells of a 2X11 cable...
                                            strUpdateCMD = "INSERT INTO ScanData (Station,BWO,STEP,RDG,[DATE],QS1,CTR,ETIME,CUR1,CUR2,VB1,VB2,VB3,VB4," +
                                                "CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24," +
                                                "BT1,BT2,BT3,BT4,BT5,BT6,CGND1,CGND2,REF,GND,FV,MSV,PSV) "
                                                + "VALUES (" + station.ToString() + ",'" +                            //station number
                                                SWO2.Trim() + "','" +                          //WorkOrderNumber
                                                slaveStepNum2.ToString("00") + "'," +                                            //StepNumber
                                                currentReading.ToString() + ",#" +                                     //ReadingNumber
                                                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  //date
                                                GlobalVars.CScanData[station].QS1.ToString() + "','" +                  //QS1
                                                GlobalVars.CScanData[station].CTR.ToString() + "','" +                  //CTR
                                                temp.TotalDays.ToString("0.00000") + "','" +                                      //time elapsed in days
                                                GlobalVars.CScanData[station].currentOne.ToString("0.0") + "','" +           //CUR1
                                                GlobalVars.CScanData[station].currentTwo.ToString("0.0") + "','" +           //CUR2
                                                GlobalVars.CScanData[slaveRow].VB1.ToString("0.00") + "','" +                  //VB1
                                                GlobalVars.CScanData[slaveRow].VB2.ToString("0.00") + "','" +                  //VB2
                                                GlobalVars.CScanData[slaveRow].VB3.ToString("0.00") + "','" +                  //VB3
                                                GlobalVars.CScanData[slaveRow].VB4.ToString("0.00") + "','" +                  //VB4
                                                GlobalVars.CScanData[slaveRow].orderedCells[12].ToString("0.000") + "','" +      //CEL01
                                                GlobalVars.CScanData[slaveRow].orderedCells[13].ToString("0.000") + "','" +      //CEL02
                                                GlobalVars.CScanData[slaveRow].orderedCells[14].ToString("0.000") + "','" +      //CEL03
                                                GlobalVars.CScanData[slaveRow].orderedCells[15].ToString("0.000") + "','" +      //CEL04
                                                GlobalVars.CScanData[slaveRow].orderedCells[16].ToString("0.000") + "','" +      //CEL05
                                                GlobalVars.CScanData[slaveRow].orderedCells[17].ToString("0.000") + "','" +      //CEL06
                                                GlobalVars.CScanData[slaveRow].orderedCells[18].ToString("0.000") + "','" +      //CEL07
                                                GlobalVars.CScanData[slaveRow].orderedCells[19].ToString("0.000") + "','" +      //CEL08
                                                GlobalVars.CScanData[slaveRow].orderedCells[20].ToString("0.000") + "','" +      //CEL09
                                                GlobalVars.CScanData[slaveRow].orderedCells[21].ToString("0.000") + "','" +      //CEL10
                                                GlobalVars.CScanData[slaveRow].orderedCells[22].ToString("0.000") + "','" +     //CEL11
                                                GlobalVars.CScanData[slaveRow].orderedCells[0].ToString("0.000") + "','" +     //CEL12
                                                GlobalVars.CScanData[slaveRow].orderedCells[1].ToString("0.000") + "','" +     //CEL13
                                                GlobalVars.CScanData[slaveRow].orderedCells[2].ToString("0.000") + "','" +     //CEL14
                                                GlobalVars.CScanData[slaveRow].orderedCells[3].ToString("0.000") + "','" +     //CEL15
                                                GlobalVars.CScanData[slaveRow].orderedCells[4].ToString("0.000") + "','" +     //CEL16
                                                GlobalVars.CScanData[slaveRow].orderedCells[5].ToString("0.000") + "','" +     //CEL17
                                                GlobalVars.CScanData[slaveRow].orderedCells[6].ToString("0.000") + "','" +     //CEL18
                                                GlobalVars.CScanData[slaveRow].orderedCells[7].ToString("0.000") + "','" +     //CEL19
                                                GlobalVars.CScanData[slaveRow].orderedCells[8].ToString("0.000") + "','" +     //CEL20
                                                GlobalVars.CScanData[slaveRow].orderedCells[9].ToString("0.000") + "','" +     //CEL21
                                                GlobalVars.CScanData[slaveRow].orderedCells[10].ToString("0.000") + "','" +     //CEL22
                                                GlobalVars.CScanData[slaveRow].orderedCells[11].ToString("0.000") + "','" +     //CEL23
                                                GlobalVars.CScanData[slaveRow].orderedCells[23].ToString("0.000") + "','" +     //CEL24
                                                GlobalVars.CScanData[station].TP1.ToString("0.0") + "','" +                  //TP1
                                                GlobalVars.CScanData[station].TP2.ToString("0.0") + "','" +                  //TP2
                                                GlobalVars.CScanData[station].TP3.ToString("0.0") + "','" +                  //TP3
                                                GlobalVars.CScanData[station].TP4.ToString("0.0") + "','" +                  //TP4
                                                GlobalVars.CScanData[station].TP5.ToString("0.0") + "','" +                  //TP5
                                                "0.0" + "','" +                                                         //TP6
                                                GlobalVars.CScanData[slaveRow].cellGND1.ToString("0.000") + "','" +             //CGND1
                                                GlobalVars.CScanData[slaveRow].cellGND2.ToString("0.000") + "','" +             //CGND2
                                                GlobalVars.CScanData[slaveRow].ref95V.ToString("0.000") + "','" +               //ref
                                                GlobalVars.CScanData[slaveRow].ch0GND.ToString("0.000") + "','" +               //GND
                                                GlobalVars.CScanData[slaveRow].plus5V.ToString("0.000") + "','" +              //FV
                                                GlobalVars.CScanData[slaveRow].minus15.ToString("0.00") + "','" +              //MSV
                                                GlobalVars.CScanData[slaveRow].plus15.ToString("0.00") +                       //PSV
                                                "');";

                                            myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                                            lock (dataBaseLock)
                                            {
                                                myAccessConn.Open();
                                                myAccessCommand.ExecuteNonQuery();
                                                myAccessConn.Close();
                                            }
                                        }
                                        else if (SWO2 != "")
                                        {
                                            // this is the case where the second work order is the middle 7 cells of a 3X7 cable...
                                            strUpdateCMD = "INSERT INTO ScanData (Station,BWO,STEP,RDG,[DATE],QS1,CTR,ETIME,CUR1,CUR2,VB1,VB2,VB3,VB4," +
                                            "CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24," +
                                            "BT1,BT2,BT3,BT4,BT5,BT6,CGND1,CGND2,REF,GND,FV,MSV,PSV) "
                                            + "VALUES (" + slaveRow.ToString() + ",'" +                            //station number
                                            SWO2.Trim() + "','" +                          //WorkOrderNumber
                                            slaveStepNum2.ToString("00") + "'," +                                            //StepNumber
                                            currentReading.ToString() + ",#" +                                     //ReadingNumber
                                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  //date
                                            GlobalVars.CScanData[slaveRow].QS1.ToString() + "','" +                  //QS1
                                            GlobalVars.CScanData[slaveRow].CTR.ToString() + "','" +                  //CTR
                                            temp.TotalDays.ToString("0.00000") + "','" +                                      //time elapsed in days
                                            GlobalVars.CScanData[station].currentOne.ToString("0.0") + "','" +           //CUR1
                                            GlobalVars.CScanData[station].currentTwo.ToString("0.0") + "','" +           //CUR2
                                            GlobalVars.CScanData[slaveRow].VB1.ToString("0.00") + "','" +                  //VB1
                                            GlobalVars.CScanData[slaveRow].VB2.ToString("0.00") + "','" +                  //VB2
                                            GlobalVars.CScanData[slaveRow].VB3.ToString("0.00") + "','" +                  //VB3
                                            GlobalVars.CScanData[slaveRow].VB4.ToString("0.00") + "','" +                  //VB4
                                            GlobalVars.CScanData[slaveRow].orderedCells[8].ToString("0.000") + "','" +      //CEL01
                                            GlobalVars.CScanData[slaveRow].orderedCells[9].ToString("0.000") + "','" +      //CEL02
                                            GlobalVars.CScanData[slaveRow].orderedCells[10].ToString("0.000") + "','" +      //CEL03
                                            GlobalVars.CScanData[slaveRow].orderedCells[11].ToString("0.000") + "','" +      //CEL04
                                            GlobalVars.CScanData[slaveRow].orderedCells[12].ToString("0.000") + "','" +      //CEL05
                                            GlobalVars.CScanData[slaveRow].orderedCells[13].ToString("0.000") + "','" +      //CEL06
                                            GlobalVars.CScanData[slaveRow].orderedCells[14].ToString("0.000") + "','" +      //CEL07
                                            GlobalVars.CScanData[slaveRow].orderedCells[15].ToString("0.000") + "','" +      //CEL08
                                            GlobalVars.CScanData[slaveRow].orderedCells[16].ToString("0.000") + "','" +      //CEL09
                                            GlobalVars.CScanData[slaveRow].orderedCells[17].ToString("0.000") + "','" +      //CEL10
                                            GlobalVars.CScanData[slaveRow].orderedCells[18].ToString("0.000") + "','" +     //CEL11
                                            GlobalVars.CScanData[slaveRow].orderedCells[19].ToString("0.000") + "','" +     //CEL12
                                            GlobalVars.CScanData[slaveRow].orderedCells[20].ToString("0.000") + "','" +     //CEL13
                                            GlobalVars.CScanData[slaveRow].orderedCells[21].ToString("0.000") + "','" +     //CEL14
                                            GlobalVars.CScanData[slaveRow].orderedCells[22].ToString("0.000") + "','" +     //CEL15
                                            GlobalVars.CScanData[slaveRow].orderedCells[23].ToString("0.000") + "','" +     //CEL16
                                            GlobalVars.CScanData[slaveRow].orderedCells[0].ToString("0.000") + "','" +     //CEL17
                                            GlobalVars.CScanData[slaveRow].orderedCells[1].ToString("0.000") + "','" +     //CEL18
                                            GlobalVars.CScanData[slaveRow].orderedCells[2].ToString("0.000") + "','" +     //CEL19
                                            GlobalVars.CScanData[slaveRow].orderedCells[3].ToString("0.000") + "','" +     //CEL20
                                            GlobalVars.CScanData[slaveRow].orderedCells[4].ToString("0.000") + "','" +     //CEL21
                                            GlobalVars.CScanData[slaveRow].orderedCells[5].ToString("0.000") + "','" +     //CEL22
                                            GlobalVars.CScanData[slaveRow].orderedCells[6].ToString("0.000") + "','" +     //CEL23
                                            GlobalVars.CScanData[slaveRow].orderedCells[7].ToString("0.000") + "','" +     //CEL24
                                            GlobalVars.CScanData[station].TP1.ToString("0.0") + "','" +                  //TP1
                                            GlobalVars.CScanData[station].TP2.ToString("0.0") + "','" +                  //TP2
                                            GlobalVars.CScanData[station].TP3.ToString("0.0") + "','" +                  //TP3
                                            GlobalVars.CScanData[station].TP4.ToString("0.0") + "','" +                  //TP4
                                            GlobalVars.CScanData[station].TP5.ToString("0.0") + "','" +                  //TP5
                                            "0.0" + "','" +                                                         //TP6
                                            GlobalVars.CScanData[slaveRow].cellGND1.ToString("0.000") + "','" +             //CGND1
                                            GlobalVars.CScanData[slaveRow].cellGND2.ToString("0.000") + "','" +             //CGND2
                                            GlobalVars.CScanData[slaveRow].ref95V.ToString("0.000") + "','" +               //ref
                                            GlobalVars.CScanData[slaveRow].ch0GND.ToString("0.000") + "','" +               //GND
                                            GlobalVars.CScanData[slaveRow].plus5V.ToString("0.000") + "','" +              //FV
                                            GlobalVars.CScanData[slaveRow].minus15.ToString("0.00") + "','" +              //MSV
                                            GlobalVars.CScanData[slaveRow].plus15.ToString("0.00") +                       //PSV
                                            "');";

                                            myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                                            lock (dataBaseLock)
                                            {
                                                myAccessConn.Open();
                                                myAccessCommand.ExecuteNonQuery();
                                                myAccessConn.Close();
                                            }
                                        }

                                        if (MWO3 != "")
                                        {
                                            // this is the case where the third work order is the last 7 cells of a 3X7 cable...
                                            strUpdateCMD = "INSERT INTO ScanData (Station,BWO,STEP,RDG,[DATE],QS1,CTR,ETIME,CUR1,CUR2,VB1,VB2,VB3,VB4," +
                                            "CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24," +
                                            "BT1,BT2,BT3,BT4,BT5,BT6,CGND1,CGND2,REF,GND,FV,MSV,PSV) "
                                            + "VALUES (" + slaveRow.ToString() + ",'" +                            //station number
                                            SWO3.Trim() + "','" +                          //WorkOrderNumber
                                            slaveStepNum3.ToString("00") + "'," +                                            //StepNumber
                                            currentReading.ToString() + ",#" +                                     //ReadingNumber
                                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  //date
                                            GlobalVars.CScanData[slaveRow].QS1.ToString() + "','" +                  //QS1
                                            GlobalVars.CScanData[slaveRow].CTR.ToString() + "','" +                  //CTR
                                            temp.TotalDays.ToString("0.00000") + "','" +                                      //time elapsed in days
                                            GlobalVars.CScanData[station].currentOne.ToString("0.0") + "','" +           //CUR1
                                            GlobalVars.CScanData[station].currentTwo.ToString("0.0") + "','" +           //CUR2
                                            GlobalVars.CScanData[slaveRow].VB1.ToString("0.00") + "','" +                  //VB1
                                            GlobalVars.CScanData[slaveRow].VB2.ToString("0.00") + "','" +                  //VB2
                                            GlobalVars.CScanData[slaveRow].VB3.ToString("0.00") + "','" +                  //VB3
                                            GlobalVars.CScanData[slaveRow].VB4.ToString("0.00") + "','" +                  //VB4
                                            GlobalVars.CScanData[slaveRow].orderedCells[16].ToString("0.000") + "','" +      //CEL01
                                            GlobalVars.CScanData[slaveRow].orderedCells[17].ToString("0.000") + "','" +      //CEL02
                                            GlobalVars.CScanData[slaveRow].orderedCells[18].ToString("0.000") + "','" +      //CEL03
                                            GlobalVars.CScanData[slaveRow].orderedCells[19].ToString("0.000") + "','" +      //CEL04
                                            GlobalVars.CScanData[slaveRow].orderedCells[20].ToString("0.000") + "','" +      //CEL05
                                            GlobalVars.CScanData[slaveRow].orderedCells[21].ToString("0.000") + "','" +      //CEL06
                                            GlobalVars.CScanData[slaveRow].orderedCells[22].ToString("0.000") + "','" +      //CEL07
                                            GlobalVars.CScanData[slaveRow].orderedCells[23].ToString("0.000") + "','" +      //CEL08
                                            GlobalVars.CScanData[slaveRow].orderedCells[0].ToString("0.000") + "','" +      //CEL09
                                            GlobalVars.CScanData[slaveRow].orderedCells[1].ToString("0.000") + "','" +      //CEL10
                                            GlobalVars.CScanData[slaveRow].orderedCells[2].ToString("0.000") + "','" +     //CEL11
                                            GlobalVars.CScanData[slaveRow].orderedCells[3].ToString("0.000") + "','" +     //CEL12
                                            GlobalVars.CScanData[slaveRow].orderedCells[4].ToString("0.000") + "','" +     //CEL13
                                            GlobalVars.CScanData[slaveRow].orderedCells[5].ToString("0.000") + "','" +     //CEL14
                                            GlobalVars.CScanData[slaveRow].orderedCells[6].ToString("0.000") + "','" +     //CEL15
                                            GlobalVars.CScanData[slaveRow].orderedCells[7].ToString("0.000") + "','" +     //CEL16
                                            GlobalVars.CScanData[slaveRow].orderedCells[8].ToString("0.000") + "','" +     //CEL17
                                            GlobalVars.CScanData[slaveRow].orderedCells[9].ToString("0.000") + "','" +     //CEL18
                                            GlobalVars.CScanData[slaveRow].orderedCells[10].ToString("0.000") + "','" +     //CEL19
                                            GlobalVars.CScanData[slaveRow].orderedCells[11].ToString("0.000") + "','" +     //CEL20
                                            GlobalVars.CScanData[slaveRow].orderedCells[12].ToString("0.000") + "','" +     //CEL21
                                            GlobalVars.CScanData[slaveRow].orderedCells[13].ToString("0.000") + "','" +     //CEL22
                                            GlobalVars.CScanData[slaveRow].orderedCells[14].ToString("0.000") + "','" +     //CEL23
                                            GlobalVars.CScanData[slaveRow].orderedCells[15].ToString("0.000") + "','" +     //CEL24
                                            GlobalVars.CScanData[station].TP1.ToString("0.0") + "','" +                  //TP1
                                            GlobalVars.CScanData[station].TP2.ToString("0.0") + "','" +                  //TP2
                                            GlobalVars.CScanData[station].TP3.ToString("0.0") + "','" +                  //TP3
                                            GlobalVars.CScanData[station].TP4.ToString("0.0") + "','" +                  //TP4
                                            GlobalVars.CScanData[station].TP5.ToString("0.0") + "','" +                  //TP5
                                            "0.0" + "','" +                                                         //TP6
                                            GlobalVars.CScanData[slaveRow].cellGND1.ToString("0.000") + "','" +             //CGND1
                                            GlobalVars.CScanData[slaveRow].cellGND2.ToString("0.000") + "','" +             //CGND2
                                            GlobalVars.CScanData[slaveRow].ref95V.ToString("0.000") + "','" +               //ref
                                            GlobalVars.CScanData[slaveRow].ch0GND.ToString("0.000") + "','" +               //GND
                                            GlobalVars.CScanData[slaveRow].plus5V.ToString("0.000") + "','" +              //FV
                                            GlobalVars.CScanData[slaveRow].minus15.ToString("0.00") + "','" +              //MSV
                                            GlobalVars.CScanData[slaveRow].plus15.ToString("0.00") +                       //PSV
                                            "');";

                                            myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                                            lock (dataBaseLock)
                                            {
                                                myAccessConn.Open();
                                                myAccessCommand.ExecuteNonQuery();
                                                myAccessConn.Close();
                                            }

                                        }  // end if

                                    } // end if

                                    //if this is the first scan we also need to rerun fill plot combos and we are still on the station...
                                    if (firstRun && station == dataGridView1.CurrentRow.Index)
                                    {
                                        oldRow = 99;
                                        fillPlotCombos(station);
                                        firstRun = false;
                                    }
                                    //else we need to add the current point to the graphMainSet dataTable if this test is running on the current row...
                                    else if (station == dataGridView1.CurrentRow.Index)
                                    {
                                        //update graphmainset
                                        DataRow newRow = graphMainSet.Tables[0].NewRow();

                                        newRow["Station"] = station.ToString();
                                        newRow["BWO"] = MWO1.Trim();
                                        newRow["STEP"] = stepNum.ToString("00");
                                        newRow["RDG"] = currentReading.ToString();
                                        newRow["DATE"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        newRow["QS1"] = GlobalVars.CScanData[station].QS1.ToString();
                                        newRow["CTR"] = GlobalVars.CScanData[station].CTR.ToString();
                                        newRow["ETIME"] = temp.TotalDays.ToString("0.00000");
                                        newRow["CUR1"] = GlobalVars.CScanData[station].currentOne.ToString("0.0");
                                        newRow["CUR2"] = GlobalVars.CScanData[station].currentTwo.ToString("0.0");
                                        newRow["VB1"] = GlobalVars.CScanData[station].VB1.ToString("0.00");
                                        newRow["VB2"] = GlobalVars.CScanData[station].VB2.ToString("0.00");
                                        newRow["VB3"] = GlobalVars.CScanData[station].VB3.ToString("0.00");
                                        newRow["VB4"] = GlobalVars.CScanData[station].VB4.ToString("0.00");
                                        newRow["CEL01"] = GlobalVars.CScanData[station].orderedCells[0].ToString("0.000");
                                        newRow["CEL02"] = GlobalVars.CScanData[station].orderedCells[1].ToString("0.000");
                                        newRow["CEL03"] = GlobalVars.CScanData[station].orderedCells[2].ToString("0.000");
                                        newRow["CEL04"] = GlobalVars.CScanData[station].orderedCells[3].ToString("0.000");
                                        newRow["CEL05"] = GlobalVars.CScanData[station].orderedCells[4].ToString("0.000");
                                        newRow["CEL06"] = GlobalVars.CScanData[station].orderedCells[5].ToString("0.000");
                                        newRow["CEL07"] = GlobalVars.CScanData[station].orderedCells[6].ToString("0.000");
                                        newRow["CEL08"] = GlobalVars.CScanData[station].orderedCells[7].ToString("0.000");
                                        newRow["CEL09"] = GlobalVars.CScanData[station].orderedCells[8].ToString("0.000");
                                        newRow["CEL10"] = GlobalVars.CScanData[station].orderedCells[9].ToString("0.000");
                                        newRow["CEL11"] = GlobalVars.CScanData[station].orderedCells[10].ToString("0.000");
                                        newRow["CEL12"] = GlobalVars.CScanData[station].orderedCells[11].ToString("0.000");
                                        newRow["CEL13"] = GlobalVars.CScanData[station].orderedCells[12].ToString("0.000");
                                        newRow["CEL14"] = GlobalVars.CScanData[station].orderedCells[13].ToString("0.000");
                                        newRow["CEL15"] = GlobalVars.CScanData[station].orderedCells[14].ToString("0.000");
                                        newRow["CEL16"] = GlobalVars.CScanData[station].orderedCells[15].ToString("0.000");
                                        newRow["CEL17"] = GlobalVars.CScanData[station].orderedCells[16].ToString("0.000");
                                        newRow["CEL18"] = GlobalVars.CScanData[station].orderedCells[17].ToString("0.000");
                                        newRow["CEL19"] = GlobalVars.CScanData[station].orderedCells[18].ToString("0.000");
                                        newRow["CEL20"] = GlobalVars.CScanData[station].orderedCells[19].ToString("0.000");
                                        newRow["CEL21"] = GlobalVars.CScanData[station].orderedCells[20].ToString("0.000");
                                        newRow["CEL22"] = GlobalVars.CScanData[station].orderedCells[21].ToString("0.000");
                                        newRow["CEL23"] = GlobalVars.CScanData[station].orderedCells[22].ToString("0.000");
                                        newRow["CEL24"] = GlobalVars.CScanData[station].orderedCells[23].ToString("0.000");
                                        newRow["BT1"] = GlobalVars.CScanData[station].TP1.ToString("0.0");
                                        newRow["BT2"] = GlobalVars.CScanData[station].TP2.ToString("0.0");
                                        newRow["BT3"] = GlobalVars.CScanData[station].TP3.ToString("0.0");
                                        newRow["BT4"] = GlobalVars.CScanData[station].TP4.ToString("0.0");
                                        newRow["BT5"] = GlobalVars.CScanData[station].TP5.ToString("0.0");
                                        newRow["BT6"] = "0.0";
                                        newRow["CGND1"] = GlobalVars.CScanData[station].cellGND1.ToString("0.000");
                                        newRow["CGND2"] = GlobalVars.CScanData[station].cellGND2.ToString("0.000");
                                        newRow["REF"] = GlobalVars.CScanData[station].ref95V.ToString("0.000");
                                        newRow["GND"] = GlobalVars.CScanData[station].ch0GND.ToString("0.000");
                                        newRow["FV"] = GlobalVars.CScanData[station].plus5V.ToString("0.000");
                                        newRow["MSV"] = GlobalVars.CScanData[station].minus15.ToString("0.00");
                                        newRow["PSV"] = GlobalVars.CScanData[station].plus15.ToString("0.00");

                                        graphMainSet.Tables[0].Rows.Add(newRow);

                                    }
                                    else if (firstRun && slaveRow == dataGridView1.CurrentRow.Index)
                                    {
                                        oldRow = 99;
                                        fillPlotCombos(slaveRow);
                                        firstRun = false;
                                    }
                                    //else we need to add the current point to the graphMainSet dataTable if this test is running on the current row...
                                    else if (slaveRow == dataGridView1.CurrentRow.Index)
                                    {
                                        //update graphmainset
                                        DataRow newRow = graphMainSet.Tables[0].NewRow();

                                        newRow["Station"] = slaveRow.ToString();
                                        newRow["BWO"] = SWO1.Trim();
                                        newRow["STEP"] = stepNum.ToString("00");
                                        newRow["RDG"] = currentReading.ToString();
                                        newRow["DATE"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        newRow["QS1"] = GlobalVars.CScanData[slaveRow].QS1.ToString();
                                        newRow["CTR"] = GlobalVars.CScanData[slaveRow].CTR.ToString();
                                        newRow["ETIME"] = temp.TotalDays.ToString("0.00000");
                                        newRow["CUR1"] = GlobalVars.CScanData[station].currentOne.ToString("0.0");
                                        newRow["CUR2"] = GlobalVars.CScanData[station].currentTwo.ToString("0.0");
                                        newRow["VB1"] = GlobalVars.CScanData[slaveRow].VB1.ToString("0.00");
                                        newRow["VB2"] = GlobalVars.CScanData[slaveRow].VB2.ToString("0.00");
                                        newRow["VB3"] = GlobalVars.CScanData[slaveRow].VB3.ToString("0.00");
                                        newRow["VB4"] = GlobalVars.CScanData[slaveRow].VB4.ToString("0.00");
                                        newRow["CEL01"] = GlobalVars.CScanData[slaveRow].orderedCells[0].ToString("0.000");
                                        newRow["CEL02"] = GlobalVars.CScanData[slaveRow].orderedCells[1].ToString("0.000");
                                        newRow["CEL03"] = GlobalVars.CScanData[slaveRow].orderedCells[2].ToString("0.000");
                                        newRow["CEL04"] = GlobalVars.CScanData[slaveRow].orderedCells[3].ToString("0.000");
                                        newRow["CEL05"] = GlobalVars.CScanData[slaveRow].orderedCells[4].ToString("0.000");
                                        newRow["CEL06"] = GlobalVars.CScanData[slaveRow].orderedCells[5].ToString("0.000");
                                        newRow["CEL07"] = GlobalVars.CScanData[slaveRow].orderedCells[6].ToString("0.000");
                                        newRow["CEL08"] = GlobalVars.CScanData[slaveRow].orderedCells[7].ToString("0.000");
                                        newRow["CEL09"] = GlobalVars.CScanData[slaveRow].orderedCells[8].ToString("0.000");
                                        newRow["CEL10"] = GlobalVars.CScanData[slaveRow].orderedCells[9].ToString("0.000");
                                        newRow["CEL11"] = GlobalVars.CScanData[slaveRow].orderedCells[10].ToString("0.000");
                                        newRow["CEL12"] = GlobalVars.CScanData[slaveRow].orderedCells[11].ToString("0.000");
                                        newRow["CEL13"] = GlobalVars.CScanData[slaveRow].orderedCells[12].ToString("0.000");
                                        newRow["CEL14"] = GlobalVars.CScanData[slaveRow].orderedCells[13].ToString("0.000");
                                        newRow["CEL15"] = GlobalVars.CScanData[slaveRow].orderedCells[14].ToString("0.000");
                                        newRow["CEL16"] = GlobalVars.CScanData[slaveRow].orderedCells[15].ToString("0.000");
                                        newRow["CEL17"] = GlobalVars.CScanData[slaveRow].orderedCells[16].ToString("0.000");
                                        newRow["CEL18"] = GlobalVars.CScanData[slaveRow].orderedCells[17].ToString("0.000");
                                        newRow["CEL19"] = GlobalVars.CScanData[slaveRow].orderedCells[18].ToString("0.000");
                                        newRow["CEL20"] = GlobalVars.CScanData[slaveRow].orderedCells[19].ToString("0.000");
                                        newRow["CEL21"] = GlobalVars.CScanData[slaveRow].orderedCells[20].ToString("0.000");
                                        newRow["CEL22"] = GlobalVars.CScanData[slaveRow].orderedCells[21].ToString("0.000");
                                        newRow["CEL23"] = GlobalVars.CScanData[slaveRow].orderedCells[22].ToString("0.000");
                                        newRow["CEL24"] = GlobalVars.CScanData[slaveRow].orderedCells[23].ToString("0.000");
                                        newRow["BT1"] = GlobalVars.CScanData[station].TP1.ToString("0.0");
                                        newRow["BT2"] = GlobalVars.CScanData[station].TP2.ToString("0.0");
                                        newRow["BT3"] = GlobalVars.CScanData[station].TP3.ToString("0.0");
                                        newRow["BT4"] = GlobalVars.CScanData[station].TP4.ToString("0.0");
                                        newRow["BT5"] = GlobalVars.CScanData[station].TP5.ToString("0.0");
                                        newRow["BT6"] = "0.0";
                                        newRow["CGND1"] = GlobalVars.CScanData[slaveRow].cellGND1.ToString("0.000");
                                        newRow["CGND2"] = GlobalVars.CScanData[slaveRow].cellGND2.ToString("0.000");
                                        newRow["REF"] = GlobalVars.CScanData[slaveRow].ref95V.ToString("0.000");
                                        newRow["GND"] = GlobalVars.CScanData[slaveRow].ch0GND.ToString("0.000");
                                        newRow["FV"] = GlobalVars.CScanData[slaveRow].plus5V.ToString("0.000");
                                        newRow["MSV"] = GlobalVars.CScanData[slaveRow].minus15.ToString("0.00");
                                        newRow["PSV"] = GlobalVars.CScanData[slaveRow].plus15.ToString("0.00");

                                        graphMainSet.Tables[0].Rows.Add(newRow);

                                    }

                                }
                                catch (Exception ex)
                                {

                                    myAccessConn.Close();
                                    updateD(station, 5, false);
                                    if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                    //clear values from d
                                    updateD(station, 7, ("Error"));
                                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                                    cRunTest[station].Cancel();
                                    //update the finish time
                                    updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        sendNote(station, 1, "Error: Failed to retrieve the required data from the DataBase.");
                                        startNewTestToolStripMenuItem.Enabled = true;
                                        resumeTestToolStripMenuItem.Enabled = true;
                                        stopTestToolStripMenuItem.Enabled = false;
                                        MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    });
                                    return;
                                }
                                #endregion

                                // finally update the reading
                                currentReading++;
                            }

                            //Now update the timer
                            eTime = stopwatch.Elapsed.Add(offset);
                            if (eTime.Days > 0)
                            {
                                eTimeS = eTime.ToString(@"dd\:hh\:mm\:ss");
                            }
                            else
                            {
                                eTimeS = eTime.ToString(@"hh\:mm\:ss");
                            }

                            if (oldETime != eTimeS)
                            {
                                try
                                {
                                    updateD(station, 6, eTimeS);
                                    if (MasterSlaveTest) { updateD(slaveRow, 6, eTimeS); }
                                }
                                catch { }

                            }
                            oldETime = eTimeS;

                            #region Here is where we are going to look for charging issues!
                            //Lets test for a charger issue now
                            // there are going to be three sections, ICA section, CCA section and SHUNT section
                            #region CScan Tests, Temp drift check and falling cells test

                            #region connections tests

                            //Lets look to see if the cables have become disconnected...
                            //start with a total Cscan test...
                            if (dataGridView1.Rows[station].Cells[4].Style.BackColor == Color.Red)
                            {
                                badCSCanReads++;
                            }
                            else
                            {
                                badCSCanReads = 0;
                            }

                            if (badCSCanReads > 50)
                            {
                                //cancel the test
                                cRunTest[station].Cancel();
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    sendNote(station, 1, "CScan communcations are failing. Please check the CScan connection and resume or restart the test.");
                                });
                            }

                            //check them
                            if (GlobalVars.CScanData[station].shuntCableType == "NONE")
                            {
                                shuntCon = false;
                            }
                            else
                            {
                                shuntCon = true;
                            }
                            if (GlobalVars.CScanData[station].tempPlateType == "NONE")
                            {
                                tempCon = false;
                            }
                            else
                            {
                                tempCon = true;
                            }
                            if (GlobalVars.CScanData[station].cellCableType == "NONE")
                            {
                                cellCon = false;
                            }
                            else
                            {
                                cellCon = true;
                            }

                            //figure out if something bad happened
                            if (shuntCon == false && oldShuntCon == true)
                            {
                                //shunt just got disconnected!
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    sendNote(station, 1, "Shunt disconnected from station " + station.ToString());
                                });
                            }
                            if (tempCon == false && oldTempCon == true)
                            {
                                //shunt just got disconnected!
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    sendNote(station, 1, "Temp-Plate disconnected from station " + station.ToString());
                                });
                            }
                            if (cellCon == false && oldCellCon == true)
                            {
                                //shunt just got disconnected!
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    sendNote(station, 1, "Cells cable disconnected from station " + station.ToString());
                                });
                            }

                            //reset the old ones...
                            oldShuntCon = shuntCon;
                            oldTempCon = tempCon;
                            oldCellCon = cellCon;

                            #endregion

                            #region temp rise test

                            //Lets also look for a temperature rise on the temp-plate if we are in charge mode....
                            //Skip for discharge
                            if (!(testType.Contains("Dis") || testType.Contains("Cap") || testType.Contains("Shorting")))
                            {
                                //also check that we don't have an ICA running in a custom discharge mode...
                                if ((d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) && (GlobalVars.ICData[Cstation].testMode.ToString().Contains("30") || GlobalVars.ICData[Cstation].testMode.ToString().Contains("31") || GlobalVars.ICData[Cstation].testMode.ToString().Contains("32")))
                                {
                                    // we have a custom ICA discharge test...
                                }
                                else
                                {
                                    double tempTemp = (GlobalVars.CScanData[station].TP1 + GlobalVars.CScanData[station].TP2 + GlobalVars.CScanData[station].TP3 + GlobalVars.CScanData[station].TP4) / 4;
                                    if ((tempTemp > (startTemp + 5)) && memTemp == 0)
                                    {
                                        // we got the first 5 C rise
                                        // warn the user
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            sendNote(station, 3, "Temperature has risen 5C on template");
                                        });
                                        //update memTemp
                                        memTemp += 1;
                                    }
                                    else if ((tempTemp > (startTemp + 10)) && memTemp == 1)
                                    {
                                        // we got the a  10 C rise
                                        // warn the user
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            sendNote(station, 2, "Temperature has risen 10C on template");
                                        });
                                        //update memTemp
                                        memTemp += 1;
                                    }
                                    else if ((tempTemp > (startTemp + 15)) && memTemp == 2)
                                    {
                                        // we got the a 15 C rise
                                        // warn the user
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            sendNote(station, 1, "Temperature has risen 15C on template");
                                        });
                                        //update memTemp
                                        memTemp += 1;
                                    }
                                    else if ((tempTemp > (startTemp + 20)) && memTemp == 3)
                                    {
                                        // we got the a 20 C rise
                                        // this is out of hand!  Lets quit the test. If we can...
                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            sendNote(station, 1, "Temperature has risen 20C on template");
                                        });
                                        // here is where we cancel the test if we can...
                                        if ((d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) || d.Rows[station][10].ToString().Contains("CCA") && !runAsShunt)
                                        {
                                            cRunTest[station].Cancel();
                                            this.Invoke((MethodInvoker)delegate()
                                            {
                                                sendNote(station, 1, "Drastic temperature rise on the battery! Cancelling Test!");
                                            });
                                        }
                                        else
                                        {
                                            this.Invoke((MethodInvoker)delegate()
                                            {
                                                sendNote(station, 1, "Drastic temperature rise on the battery! Cancel the test!");
                                            });
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region failing cell test
                            //we are going to look for failing cells here...

                            //only check once every 20 secs...

                            //if we have the cell check turned on and the we are in charge check the cells for negative slopes...
                            if (DateTime.Now.Ticks - cellCheckTimer.Ticks > (GlobalVars.DCVPeriod * 10000000) && GlobalVars.DecliningCellVoltageTestEnabled && !(testType.Contains("Dis") || testType.Contains("Cap") || testType.Contains("Shorting")))
                            {

                                //also check that we don't have an ICA running in a custom discharge mode...
                                if ((d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) && (GlobalVars.ICData[Cstation].testMode.ToString().Contains("30") || GlobalVars.ICData[Cstation].testMode.ToString().Contains("31") || GlobalVars.ICData[Cstation].testMode.ToString().Contains("32")))
                                {
                                    // we have a custom ICA discharge test...
                                    // skip this step...
                                }
                                else
                                {
                                    cellCheckTimer = DateTime.Now;

                                    int Cells;

                                    if ((int)pci.Rows[station][3] == -1)
                                    {
                                        Cells = GlobalVars.CScanData[station].cellsToDisplay;
                                    }
                                    else
                                    {
                                        Cells = (int)pci.Rows[station][3];
                                    }

                                    // move the old cells back a place
                                    for (int i = 1; i < 4; i++)
                                    {
                                        for (int j = 0; j < 24; j++)
                                        {
                                            cellHistory[j, i - 1] = cellHistory[j, i];
                                        }
                                    }
                                    // now load in the new values
                                    for (int j = 0; j < 24; j++)
                                    {
                                        cellHistory[j, 3] = GlobalVars.CScanData[station].orderedCells[j];
                                    }

                                    //finally test the cells for three consecutive negative slopes..
                                    for (int j = 0; j < Cells; j++)
                                    {
                                        if ((cellHistory[j, 1] - cellHistory[j, 0]) < (double)(GlobalVars.DecliningCellVoltageThres) && (cellHistory[j, 2] - cellHistory[j, 1]) < (double)(GlobalVars.DecliningCellVoltageThres) && (cellHistory[j, 3] - cellHistory[j, 2]) < (double)(GlobalVars.DecliningCellVoltageThres))
                                        {
                                            //we need to quit...
                                            // here is where we cancel the test if we can...
                                            if ((d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) || d.Rows[station][10].ToString().Contains("CCA") && !runAsShunt)
                                            {
                                                cRunTest[station].Cancel();
                                                this.Invoke((MethodInvoker)delegate()
                                                {
                                                    sendNote(station, 1, "Cell " + (GlobalVars.Pos2Neg ? (Cells - j) : (j + 1)).ToString() + " voltage is reversing! Cancelling Test!");
                                                });
                                            }
                                            else
                                            {
                                                this.Invoke((MethodInvoker)delegate()
                                                {
                                                    sendNote(station, 1, "Cell " + (GlobalVars.Pos2Neg ? (Cells - j) : (j + 1)).ToString() + " voltage is reversing! Cancel the test!");
                                                });
                                            }

                                        }
                                    }

                                    //do we need to do the same for the slave?
                                    if (MasterSlaveTest)
                                    {
                                        if ((int)pci.Rows[slaveRow][3] == -1)
                                        {
                                            Cells = GlobalVars.CScanData[slaveRow].cellsToDisplay;
                                        }
                                        else
                                        {
                                            Cells = (int)pci.Rows[slaveRow][3];
                                        }

                                        // move the old cells back a place
                                        for (int i = 1; i < 4; i++)
                                        {
                                            for (int j = 0; j < 24; j++)
                                            {
                                                slaveCellHistory[j, i - 1] = slaveCellHistory[j, i];
                                            }
                                        }
                                        // now load in the new values
                                        for (int j = 0; j < 24; j++)
                                        {
                                            slaveCellHistory[j, 3] = GlobalVars.CScanData[slaveRow].orderedCells[j];
                                        }

                                        //finally test the cells for three consecutive negative slopes..
                                        for (int j = 0; j < Cells; j++)
                                        {
                                            if ((slaveCellHistory[j, 1] - slaveCellHistory[j, 0]) < (double)(GlobalVars.DecliningCellVoltageThres) && (slaveCellHistory[j, 2] - slaveCellHistory[j, 1]) < (double)(GlobalVars.DecliningCellVoltageThres) && (slaveCellHistory[j, 3] - slaveCellHistory[j, 2]) < (double)(GlobalVars.DecliningCellVoltageThres))
                                            {
                                                //we need to quit...
                                                // here is where we cancel the test if we can...
                                                if ((d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) || d.Rows[slaveRow][10].ToString().Contains("CCA") && !runAsShunt)
                                                {
                                                    cRunTest[slaveRow].Cancel();
                                                    this.Invoke((MethodInvoker)delegate()
                                                    {
                                                        sendNote(slaveRow, 1, "Cell " + (GlobalVars.Pos2Neg ? (Cells - j) : (j + 1)).ToString() + " on slave C-Scan voltage is reversing! Cancelling Test!");
                                                    });
                                                }
                                                else
                                                {
                                                    this.Invoke((MethodInvoker)delegate()
                                                    {
                                                        sendNote(slaveRow, 1, "Cell " + (GlobalVars.Pos2Neg ? (Cells - j) : (j + 1)).ToString() + "  on slave C-Scan voltage is reversing! Cancel the test!");
                                                    });
                                                }

                                            }
                                        }
                                    }
                                }
                            }

                            #endregion

                            #endregion
                            //charger specific tests

                            // No need to check for "As Received" because of the start up delay
                            #region IC checks
                            if ((d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) && !startUpDelay && !runAsShunt && !testType.Contains("Shorting"))
                            {
                                //current check part...
                                //if we have a mini that is charging...
                                if (!GlobalVars.allowZeroTest)
                                {
                                    if (d.Rows[station][10].ToString().Contains("mini") && !(testType.Contains("Cap") || testType.Contains("Discharge") || testType.Contains("Shorting")))
                                    {
                                        //badCurCount looks at current two...
                                        // for the mini case
                                        if (Math.Abs(GlobalVars.CScanData[station].currentTwo) < 0.005)
                                        {
                                            badCurCount++;
                                        }
                                        else
                                        {
                                            badCurCount = 0;
                                        }
                                    }
                                    else
                                    {
                                        // all other cases
                                        // for the mini case
                                        if (Math.Abs(GlobalVars.CScanData[station].currentOne) < 0.05)
                                        {
                                            badCurCount++;
                                        }
                                        else
                                        {
                                            badCurCount = 0;
                                        }
                                    }
                                }
                                else
                                {
                                    badCurCount = 0;
                                }


                                // this is the current check!
                                if (badCurCount > 100)
                                {
                                    // end the test!
                                    //clear values from d

                                    //also tell the charger to stop
                                    // now we need to stop the charger
                                    updateD(station, 7, "Telling Charger to Stop");
                                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Telling Charger to Stop"); }

                                    for (int j = 0; j < 5; j++)
                                    {
                                        // set KE1 to 2 ("command")
                                        GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                                        // set KE3 to stop
                                        GlobalVars.ICSettings[Cstation].KE3 = (byte)2;
                                        //Update the output string value
                                        GlobalVars.ICSettings[Cstation].UpdateOutText();

                                        for (int i = 0; i < 5; i++)
                                        {
                                            criticalNum[Cstation] = true;
                                            Thread.Sleep(1000);
                                            if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END" || d.Rows[station][11].ToString() == "RESET")
                                            {
                                                break;
                                            }

                                        }

                                        if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END" || d.Rows[station][11].ToString() == "RESET")
                                        {
                                            // set KE1 to 0 ("query")
                                            GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                                            //Update the output string value
                                            GlobalVars.ICSettings[Cstation].UpdateOutText();
                                            break;
                                        }

                                    }

                                    updateD(station, 7, ("FAILED ON " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                                    if (MasterSlaveTest) { updateD(slaveRow, 7, ("FAILED ON " + (currentReading - 1).ToString() + " of " + readings.ToString())); }
                                    updateD(station, 5, false);
                                    if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                    cRunTest[station].Cancel();
                                    //update the finish time
                                    updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);

                                    //update the gui
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        sendNote(station, 1, "Test failed. Charger is not producing any current.  Please check the charger settings.");
                                        startNewTestToolStripMenuItem.Enabled = true;
                                        resumeTestToolStripMenuItem.Enabled = true;
                                        stopTestToolStripMenuItem.Enabled = false;
                                        MessageBox.Show(this, "Test failed. Charger is not producing any current.  Please check the charger settings.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    });

                                    return;
                                }


                                ////////status checks!

                                if ((currentReading == readings || currentReading == readings + 1) && d.Rows[station][11].ToString() == "END")
                                {

                                    // we are on the last reading.  Look out for the "END"...
                                    // At the moment the test will just continue until the time is up...
                                    // doesn't seem to be an issue because all of the chargers are a little slow...

                                }
                                else if ((string)d.Rows[station][11] != "RUN" && testType != "As Received" && !testType.Contains("Shorting"))
                                {
                                    //make sure the charger has priority
                                    criticalNum[Cstation] = true;

                                    //try it a couple more times
                                    for (byte b = 0; b < 3; b++)
                                    {
                                        Thread.Sleep(2000);
                                        if ((string)d.Rows[station][11] == "RUN") { break; }
                                    }

                                    //retest the "RUN"
                                    if ((string)d.Rows[station][11] != "RUN")
                                    {
                                        updateD(station, 7, "Found Fault!");
                                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Found Fault!"); }
                                        // we got an issue!
                                        // stop the clock!
                                        stopwatch.Stop();
                                        Thread.Sleep(5000);

                                        if (d.Rows[station][11].ToString().Contains("Power") || (string)d.Rows[station][11] == "HOLD")
                                        {

                                            updateD(station, 7, "Waiting For Charger");
                                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Waiting For Charger"); }
                                            // lets put things on pause and wait for the charger to come back
                                            while ((string)d.Rows[station][11] != "HOLD" && (string)d.Rows[station][11] != "RUN")
                                            {
                                                //check for a cancel or offline!
                                                if (token.IsCancellationRequested || d.Rows[station][11].ToString() == "offline!")
                                                {

                                                    //clear values from d
                                                    updateD(station, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                                                    if (MasterSlaveTest) { updateD(slaveRow, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString())); }
                                                    updateD(station, 5, false);
                                                    if (MasterSlaveTest) { updateD(slaveRow, 5, false); }


                                                    cRunTest[station].Cancel();
                                                    //update the finish time
                                                    updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                                    //update the gui
                                                    this.Invoke((MethodInvoker)delegate()
                                                    {
                                                        sendNote(station, 2, "Test Cancelled");
                                                        startNewTestToolStripMenuItem.Enabled = true;
                                                        resumeTestToolStripMenuItem.Enabled = true;
                                                        stopTestToolStripMenuItem.Enabled = false;
                                                    });

                                                    return;
                                                }

                                                Thread.Sleep(400);
                                            }
                                            // were back!
                                            //start the charger back up!

                                            criticalNum[Cstation] = true;
                                            updateD(station, 7, "Telling Charger to Run");
                                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Telling Charger to Run"); }
                                            stopwatch.Start();

                                            ThreadPool.QueueUserWorkItem(t =>
                                            {
                                                for (int j = 0; j < 10; j++)
                                                {
                                                    // set KE1 to 2 ("command")
                                                    GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                                                    // set KE3 to run
                                                    GlobalVars.ICSettings[Cstation].KE3 = (byte)1;
                                                    //Update the output string value
                                                    GlobalVars.ICSettings[Cstation].UpdateOutText();
                                                    criticalNum[Cstation] = true;
                                                    Thread.Sleep(3000);
                                                    // set KE1 to 1 ("query")
                                                    GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                                                    //Update the output string value
                                                    GlobalVars.ICSettings[Cstation].UpdateOutText();
                                                    for (int i = 0; i < 3; i++)
                                                    {
                                                        criticalNum[Cstation] = true;
                                                        Thread.Sleep(1000);
                                                        if (d.Rows[station][11].ToString() == "RUN")
                                                        {
                                                            break;
                                                        }
                                                    }

                                                    //make sure the charger has priority
                                                    if (d.Rows[station][11].ToString() == "RUN")
                                                    {
                                                        break;
                                                    }
                                                }

                                            });                     // end thread

                                            // lay off the tests for 10 seconds...
                                            startUpDelay = true;
                                            ThreadPool.QueueUserWorkItem(t =>
                                            {
                                                Thread.Sleep(10000);
                                                startUpDelay = false;
                                            });                     // end thread


                                            Thread.Sleep(2000);
                                            updateD(station, 7, ("Reading " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                                            if (MasterSlaveTest) { updateD(slaveRow, 7, ("Reading " + (currentReading - 1).ToString() + " of " + readings.ToString())); }

                                        }// end power fail if
                                        else
                                        {
                                            //make sure the out string says to stop...

                                            // end the test!
                                            //clear values from d

                                            updateD(station, 7, ("FAILED ON " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                                            if (MasterSlaveTest) { updateD(slaveRow, 7, ("FAILED ON " + (currentReading - 1).ToString() + " of " + readings.ToString())); }
                                            updateD(station, 5, false);
                                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                                            //update the gui
                                            cRunTest[station].Cancel();
                                            //update the finish time
                                            updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                            this.Invoke((MethodInvoker)delegate()
                                            {
                                                sendNote(station, 1, "Test failed. Charger Status:  " + d.Rows[station][11].ToString());
                                                startNewTestToolStripMenuItem.Enabled = true;
                                                resumeTestToolStripMenuItem.Enabled = true;
                                                stopTestToolStripMenuItem.Enabled = false;
                                                if (!(d.Rows[station][2].ToString().Contains("Combo") && d.Rows[station][2].ToString().Contains("Discharge")) && !GlobalVars.advance2Short && d.Rows[station][11].ToString().Contains("Current Fault"))
                                                {
                                                    MessageBox.Show(this, "Test failed. Charger Status:  " + d.Rows[station][11].ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                }
                                            });
                                            return;
                                        }// end test fail else

                                    }  // end still no run if
                                } // end no run if

                                // finally let take a look at the current and voltages and the like and see if they are where they should be...
                                // this is a double check to make sure the autoconfig is behaving...
                                #region settings checks for autoConfigged ICs

                                if (GlobalVars.autoConfig && (bool)d.Rows[station][12])
                                {
                                    /* we need to check that the charger is doing what it should be...
                                    timeSet1
                                    curSet1
                                    vSet1
                                    timeSet2
                                    curSet2
                                    vSet2
                                    ohmSet
                                     * */

                                    //Debug.Print("Current:" + GlobalVars.ICData[Cstation].battCurrent.ToString());
                                    //Debug.Print("Voltage:" + GlobalVars.ICData[Cstation].battVoltage.ToString());
                                    // we are past the startup dealy, so lets see if we are reasonably close to one of the currents
                                    //check the currents
                                    if (GlobalVars.ICData[Cstation].testMode.Contains("32"))
                                    {
                                        // resistance discharge
                                        if (GlobalVars.ICData[Cstation].battVoltage < (vSet1 - 1))
                                        {
                                            badReadingCount++;
                                        }
                                    }
                                    else if (GlobalVars.ICData[Cstation].testMode.Contains("31"))
                                    {
                                        if (Math.Abs(-1 * GlobalVars.CScanData[station].currentOne - curSet1) > (0.1 + curSet1 * 0.05) && Math.Abs(-1 * GlobalVars.CScanData[station].currentOne - curSet2) > (0.1 + curSet2 * 0.05))
                                        {
                                            badReadingCount++;
                                        }
                                        else if (GlobalVars.ICData[Cstation].battVoltage < (vSet1 - 1))
                                        {
                                            badReadingCount++;
                                        }
                                        else
                                        {
                                            badReadingCount = 0;
                                        }
                                    }
                                    else if (GlobalVars.ICData[Cstation].testMode.Contains("30"))
                                    {
                                        // only test is you are over for discharge...
                                        if ((-1 * GlobalVars.CScanData[station].currentOne - curSet1) > (0.1 + curSet1 * 0.05) && (-1 * GlobalVars.CScanData[station].currentOne - curSet2) > (0.1 + curSet2 * 0.05))
                                        {
                                            badReadingCount++;
                                        }
                                        else if (GlobalVars.ICData[Cstation].battVoltage < (vSet1 - 1))
                                        {
                                            badReadingCount++;
                                        }
                                        else
                                        {
                                            badReadingCount = 0;
                                        }
                                    }
                                    else if (Math.Abs(GlobalVars.CScanData[station].currentOne - curSet1) > (0.1 + curSet1 * 0.05) && Math.Abs(GlobalVars.CScanData[station].currentOne - curSet2) > (0.1 + curSet2 * 0.05))
                                    {
                                        badReadingCount++;
                                    }
                                    else if (GlobalVars.ICData[Cstation].battVoltage > (vSet1 + 0.5) && GlobalVars.ICData[Cstation].battVoltage > (vSet2 + 0.5))
                                    {
                                        badReadingCount++;
                                    }
                                    else
                                    {
                                        badReadingCount = 0;
                                    }

                                }

                                #region cancel because of out of programmed settings error
                                if (badReadingCount > 100)
                                {
                                    // quit the test and let them know about it...
                                    // end the test!
                                    //clear values from d

                                    //update the gui

                                    //also tell the charger to stop
                                    // now we need to stop the charger
                                    updateD(station, 7, "Telling Charger to Stop");
                                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Telling Charger to Stop"); }

                                    for (int j = 0; j < 5; j++)
                                    {
                                        // set KE1 to 2 ("command")
                                        GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                                        // set KE3 to stop
                                        GlobalVars.ICSettings[Cstation].KE3 = (byte)2;
                                        //Update the output string value
                                        GlobalVars.ICSettings[Cstation].UpdateOutText();

                                        for (int i = 0; i < 5; i++)
                                        {
                                            criticalNum[Cstation] = true;
                                            Thread.Sleep(1000);
                                            if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END" || d.Rows[station][11].ToString() == "RESET")
                                            {
                                                break;
                                            }

                                        }

                                        if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END" || d.Rows[station][11].ToString() == "RESET")
                                        {
                                            // set KE1 to 0 ("query")
                                            GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                                            //Update the output string value
                                            GlobalVars.ICSettings[Cstation].UpdateOutText();
                                            break;
                                        }

                                    }

                                    updateD(station, 7, ("FAILED ON " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                                    if (MasterSlaveTest) { updateD(slaveRow, 7, ("FAILED ON " + (currentReading - 1).ToString() + " of " + readings.ToString())); }
                                    updateD(station, 5, false);
                                    if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                    cRunTest[station].Cancel();
                                    //update the finish time
                                    updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        sendNote(station, 1, "Test failed. Charger is not running as programmed.  Please check the charger comms.");
                                        startNewTestToolStripMenuItem.Enabled = true;
                                        resumeTestToolStripMenuItem.Enabled = true;
                                        stopTestToolStripMenuItem.Enabled = false;
                                        MessageBox.Show(this, "Test failed. Charger is not running as programmed.  Please check the charger comms.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    });
                                    return;

                                }
                                #endregion
                                #endregion
                            }// end IC block
                            #endregion
                            #region CCA tests
                            else if (d.Rows[station][10].ToString().Contains("CCA") && !startUpDelay && !runAsShunt && !testType.Contains("Shorting"))
                            {
                                // check that the C-Scan is still running...

                                // this is the power fail case...
                                if (d.Rows[station][11].ToString().Contains("Power"))
                                {
                                    stopwatch.Stop();
                                    //set to hold
                                    GlobalVars.cHold[station] = true;
                                    //wait for a return of power..
                                    updateD(station, 7, "Waiting For Charger");
                                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Waiting For Charger"); }

                                    while (d.Rows[station][11].ToString() != "HOLD")
                                    {


                                        // check for a cancel
                                        if (token.IsCancellationRequested)
                                        {
                                            //clear values from d
                                            updateD(station, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                                            if (MasterSlaveTest) { updateD(slaveRow, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString())); }
                                            updateD(station, 5, false);
                                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                                            //update the gui

                                            cRunTest[station].Cancel();
                                            //update the finish time
                                            updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                            this.Invoke((MethodInvoker)delegate()
                                            {
                                                sendNote(station, 2, "Test canceled while in a power fail.");
                                                startNewTestToolStripMenuItem.Enabled = true;
                                                resumeTestToolStripMenuItem.Enabled = true;
                                                stopTestToolStripMenuItem.Enabled = false;
                                                MessageBox.Show(this, "Test canceled while in a power fail.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            });
                                            return;
                                        }
                                        // wait a little before testing again...
                                        Thread.Sleep(200);
                                    }// end power fail wait

                                    //reset readings
                                    updateD(station, 7, "Restarting Charger");
                                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Restarting Charger"); }
                                    Thread.Sleep(1500);
                                    stopwatch.Start();
                                    GlobalVars.cHold[station] = false;
                                    Thread.Sleep(500);
                                    updateD(station, 7, ("Reading " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                                    if (MasterSlaveTest) { updateD(slaveRow, 7, ("Reading " + (currentReading - 1).ToString() + " of " + readings.ToString())); }

                                }

                                else if (Math.Abs(GlobalVars.CScanData[station].currentOne) < 0.2)
                                {
                                    int count = 0;
                                    while (Math.Abs(GlobalVars.CScanData[station].currentOne) < 0.2)
                                    {
                                        count++;
                                        Thread.Sleep(100);
                                        if (count > 20)
                                        {
                                            //make sure the C-Scan is back on hold
                                            GlobalVars.cHold[station] = true;

                                            //clear values from d
                                            updateD(station, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                                            if (MasterSlaveTest) { updateD(slaveRow, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString())); }
                                            updateD(station, 5, false);
                                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                                            //update the gui

                                            cRunTest[station].Cancel();
                                            //update the finish time
                                            updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                            this.Invoke((MethodInvoker)delegate()
                                            {
                                                sendNote(station, 1, "Current has fallen below minimum threshold. Please check the shunt connection and resume or restart the test.");
                                                startNewTestToolStripMenuItem.Enabled = true;
                                                resumeTestToolStripMenuItem.Enabled = true;
                                                stopTestToolStripMenuItem.Enabled = false;
                                                MessageBox.Show(this, "Current has fallen below minimum threshold. Please check the shunt connection and resume or restart the test.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            });
                                            return;
                                        }
                                    }
                                }


                            }
                            #endregion
                            #region Shunt Checks
                            else if ((d.Rows[station][10].ToString().Contains("Shunt") || runAsShunt) && !startUpDelay && !testType.Contains("Shorting"))
                            {
                                //test that the current is still above the 0.2 threshold...
                                // check that the C-Scan is still running...
                                if (Math.Abs(GlobalVars.CScanData[station].currentOne) < 0.2 && !GlobalVars.allowZeroShunt)
                                {
                                    int count = 0;

                                    while (Math.Abs(GlobalVars.CScanData[station].currentOne) < 0.2)
                                    {
                                        count++;
                                        Thread.Sleep(100);
                                        if (count > 20)
                                        {
                                            //clear values from d
                                            updateD(station, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                                            if (MasterSlaveTest) { updateD(slaveRow, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString())); }
                                            updateD(station, 5, false);
                                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }


                                            cRunTest[station].Cancel();
                                            //update the finish time
                                            updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                            //update the gui
                                            this.Invoke((MethodInvoker)delegate()
                                            {
                                                sendNote(station, 1, "Current has fallen below minimum threshold.  Please check the shunt connection and resume or restart the test.");
                                                startNewTestToolStripMenuItem.Enabled = true;
                                                resumeTestToolStripMenuItem.Enabled = true;
                                                stopTestToolStripMenuItem.Enabled = false;
                                                MessageBox.Show(this, "Current has fallen below minimum threshold. Please check the shunt connection and resume or restart the test.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            });
                                            return;
                                        }
                                    }
                                }

                            }
                            #endregion
                            #region As Received and No Charger Tests
                            else if (testType.Contains("As Received") || testType.Contains("Shorting"))
                            {
                                // no test needed for now...
                            }
                            else if (!startUpDelay)
                            {
                                // we don't have a charger or a shunt anymore!!!!
                                // stop the test!

                                //make sure the C-Scan is back on hold
                                GlobalVars.cHold[station] = true;

                                //clear values from d
                                updateD(station, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                                if (MasterSlaveTest) { updateD(slaveRow, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString())); }
                                updateD(station, 5, false);
                                if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                                //update the gui

                                cRunTest[station].Cancel();
                                //update the finish time
                                updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    sendNote(station, 1, "Charger disconnected.  Please check connection and resume or restart the test.");
                                    startNewTestToolStripMenuItem.Enabled = true;
                                    resumeTestToolStripMenuItem.Enabled = true;
                                    stopTestToolStripMenuItem.Enabled = false;
                                    MessageBox.Show(this, "Charger disconnected.  Please check connection and resume or restart the test.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                });
                                return;

                            }
                            #endregion


                            #endregion

                            //Now we should check for a cancel
                            #region cancel block
                            if (token.IsCancellationRequested)
                            {

                                if (testType == "As Received" || testType.Contains("Shorting"))
                                {
                                    //nothing to do here...
                                }
                                else if ((d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) && !runAsShunt)
                                {
                                    // now we need to stop the charger
                                    updateD(station, 7, "Telling Charger to Stop");
                                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Telling Charger to Stop"); }

                                    for (int j = 0; j < 3; j++)
                                    {
                                        // set KE1 to 2 ("command")
                                        GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                                        // set KE3 to stop
                                        GlobalVars.ICSettings[Cstation].KE3 = (byte)2;
                                        //Update the output string value
                                        GlobalVars.ICSettings[Cstation].UpdateOutText();

                                        for (int i = 0; i < 7; i++)
                                        {
                                            criticalNum[Cstation] = true;
                                            Thread.Sleep(1000);
                                            if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END" || d.Rows[station][11].ToString() == "RESET")
                                            {
                                                break;
                                            }

                                        }

                                        if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END" || d.Rows[station][11].ToString() == "RESET")
                                        {
                                            // set KE1 to 0 ("query")
                                            GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                                            //Update the output string value
                                            GlobalVars.ICSettings[Cstation].UpdateOutText();
                                            break;
                                        }

                                    }

                                }
                                else if (d.Rows[station][10].ToString().Contains("CCA") && !runAsShunt)
                                {
                                    // Put the charger back on hold...
                                    GlobalVars.cHold[station] = true;
                                }
                                else if (d.Rows[station][10].ToString().Contains("Shunt") || runAsShunt)
                                {
                                    // Also nothing to do...
                                }

                                //clear values from d
                                updateD(station, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                                if (MasterSlaveTest) { updateD(slaveRow, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString())); }
                                // only clear the test checkbox if you are not doing a combo test

                                updateD(station, 5, false);
                                if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                cRunTest[station].Cancel();
                                //update the gui
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    sendNote(station, 2, "Cancelled");
                                    startNewTestToolStripMenuItem.Enabled = true;
                                    resumeTestToolStripMenuItem.Enabled = true;
                                    stopTestToolStripMenuItem.Enabled = false;
                                });

                                //update the finish time
                                updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);
                                return;
                            }
                            #endregion

                            //every interval is defined in seconds to be safe, we'll test if we are at the correct interval every 200ms
                            Thread.Sleep(200);
                        }
                        catch (Exception ex)
                        {
                            if (ex is IndexOutOfRangeException)
                            {
                                // wait a half second and then try again...
                                //this.Invoke((MethodInvoker)delegate()
                                //    {
                                //        sendNote(station, 3, "--------");
                                //    });
                                Thread.Sleep(100);
                            }
                            else
                            {
                                MessageBox.Show(this, "Error in main test loop try:  " + ex.Message + Environment.NewLine + ex.TargetSite + Environment.NewLine + ex.HResult + Environment.NewLine + ex.StackTrace, "Main Test Loop Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                        }

                    } // end main test loop...
                    #endregion

                    ////////////////////////////////END TEST SECTION////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    #region END TEST
                    //First lets save the state of the cells at finish to use in the post test section
                    CScanDataStore lastReading = GlobalVars.CScanData[station];
                    CScanDataStore lastSlaveReading = lastReading;
                    if (slaveRow != -1)
                    {
                        lastSlaveReading = GlobalVars.CScanData[slaveRow];
                    }

                    // Ok, now let's clearn up!
                    // If we are running the charger tell it to stop and reset
                    if (testType == "As Received" || testType.Contains("Shorting"))
                    {
                        // nothing to do...
                    }
                    else if ((string)d.Rows[station][9] != "" && (d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) && !runAsShunt)
                    {
                        // now we need to stop the charger
                        updateD(station, 7, "Telling Charger to Stop");
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Telling Charger to Stop"); }
                        for (int j = 0; j < 3; j++)
                        {
                            // set KE1 to 2 ("command")
                            GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                            // set KE3 to stop
                            GlobalVars.ICSettings[Cstation].KE3 = (byte)2;
                            //Update the output string value
                            GlobalVars.ICSettings[Cstation].UpdateOutText();

                            for (int i = 0; i < 7; i++)
                            {
                                criticalNum[Cstation] = true;
                                Thread.Sleep(1000);
                                if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END" || d.Rows[station][11].ToString() == "RESET")
                                {
                                    break;
                                }

                            }

                            if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END" || d.Rows[station][11].ToString() == "RESET")
                            {
                                // set KE1 to 0 ("query")
                                GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                                //Update the output string value
                                GlobalVars.ICSettings[Cstation].UpdateOutText();
                                break;
                            }

                        }
                        //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                        // now we need to reset the charger
                        updateD(station, 7, "Resetting Charger");
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Resetting Charger"); }

                        for (int j = 0; j < 5; j++)
                        {

                            // set KE1 to 2 ("command")
                            GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                            // set KE3 to RESET
                            GlobalVars.ICSettings[Cstation].KE3 = (byte)3;
                            //Update the output string value
                            GlobalVars.ICSettings[Cstation].UpdateOutText();
                            criticalNum[Cstation] = true;
                            //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                            Thread.Sleep(5000);

                            // set KE1 to 1 ("query")
                            GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                            //Update the output string value
                            GlobalVars.ICSettings[Cstation].UpdateOutText();
                            for (int i = 0; i < 5; i++)
                            {
                                criticalNum[Cstation] = true;
                                Thread.Sleep(1000);
                                if (d.Rows[station][11].ToString() == "RESET")
                                {
                                    break;
                                }

                            }
                            if (d.Rows[station][11].ToString() == "RESET")
                            {
                                break;
                            }
                        }
                    }
                    else if (d.Rows[station][10].ToString().Contains("CCA") && !runAsShunt)
                    {
                        // Put the charger back on hold after a little while (charger clocks are typically slow)
                        ThreadPool.QueueUserWorkItem(t =>
                        {
                            Thread.Sleep(20000);
                            GlobalVars.cHold[station] = true;
                        });
                    }
                    else if (d.Rows[station][10].ToString().Contains("Shunt") || runAsShunt)
                    {
                        // Also nothing to do...
                    }


                    //update the gui
                    this.Invoke((MethodInvoker)delegate()
                    {
                        sendNote(station, 2, "Test Complete");
                        startNewTestToolStripMenuItem.Enabled = true;
                        resumeTestToolStripMenuItem.Enabled = false;
                        stopTestToolStripMenuItem.Enabled = false;
                    });

                    //Test is finished!
                    updateD(station, 6, "");
                    if (MasterSlaveTest) { updateD(slaveRow, 6, ""); }
                    updateD(station, 7, "Complete");
                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Complete"); }
                    if (!d.Rows[station][2].ToString().Contains("Combo"))
                    {
                        updateD(station, 5, false);
                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                    }

                    //update the finish time
                    updateFinishTime(MWO1, MWO2, MWO3, SWO1, SWO2, SWO3, stepNum, stepNum2, stepNum3, slaveStepNum, slaveStepNum2, slaveStepNum3);

                    // finally, we need to set the cancel token to cancelled
                    cRunTest[station].Cancel();
                    #endregion

                    ////////////////////////////////POST TEST CHECKS SECTION////////////////////////////////////////////////////////////////////////////////////////////////////
                    #region POST TEST CHECKS

                    if (GlobalVars.CapTestVarEnable && testType.Contains("Capacity"))
                    {
                        //start with the master
                        int Cells;
                        double average = 0;
                        double cleanAverage = 0;

                        //How many cells do we need to look at?
                        if ((int)pci.Rows[station][3] == -1)
                        {
                            Cells = lastReading.cellsToDisplay;
                        }
                        else
                        {
                            Cells = (int)pci.Rows[station][3];
                        }

                        // get an initial average
                        for (int i = 0; i < Cells; i++)
                        {
                            average += lastReading.orderedCells[i];
                        }
                        average /= Cells;

                        int goodCells = 0;
                        //now recalculate the average after throwing out the outliers..
                        for (int i = 0; i < Cells; i++)
                        {
                            if (Math.Abs(lastReading.orderedCells[i] - average) < (double)GlobalVars.CapTestVarValue)
                            {
                                cleanAverage += lastReading.orderedCells[i];
                                goodCells++;
                            }
                        }
                        if (goodCells > 1)
                        {
                            cleanAverage /= goodCells;
                        }

                        // let me know about it!

                        string tech = pci.Rows[station][1].ToString();
                        string test_type = d.Rows[station][2].ToString();
                        float cellLim = (float)pci.Rows[station][7];
                        string d_data = d.Rows[station][7].ToString();

                        this.Invoke((MethodInvoker)delegate()
                        {
                            //MessageBox.Show(this, "Cells out of balance:  " + average.ToString() + "  " + cleanAverage.ToString() +" (" + goodCells.ToString() + ")", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Cap_Var_Form cvf = new Cap_Var_Form(Cells, lastReading, tech, test_type, cellLim, d_data, average, MWO1, stepNum.ToString());
                            cvf.Owner = this;
                            cvf.Show();
                        });


                        if (slaveRow != -1)
                        {
                            //start with the master
                            average = 0;
                            cleanAverage = 0;

                            //How many cells do we need to look at?
                            if ((int)pci.Rows[slaveRow][3] == -1)
                            {
                                Cells = lastSlaveReading.cellsToDisplay;
                            }
                            else
                            {
                                Cells = (int)pci.Rows[slaveRow][3];
                            }

                            // get an initial average
                            for (int i = 0; i < Cells; i++)
                            {
                                average += lastSlaveReading.orderedCells[i];
                            }
                            average /= Cells;

                            goodCells = 0;
                            //now recalculate the average after throwing out the outliers..
                            for (int i = 0; i < Cells; i++)
                            {
                                if (Math.Abs(lastSlaveReading.orderedCells[i] - average) < (double)GlobalVars.CapTestVarValue)
                                {
                                    cleanAverage += lastSlaveReading.orderedCells[i];
                                    goodCells++;
                                }
                            }
                            if (goodCells > 1)
                            {
                                cleanAverage /= goodCells;
                            }

                            // let me know about it!

                            tech = pci.Rows[slaveRow][1].ToString();
                            test_type = d.Rows[slaveRow][2].ToString();
                            cellLim = (float)pci.Rows[slaveRow][7];
                            d_data = d.Rows[slaveRow][7].ToString();

                            this.Invoke((MethodInvoker)delegate()
                            {
                                //MessageBox.Show(this, "Cells out of balance:  " + average.ToString() + "  " + cleanAverage.ToString() +" (" + goodCells.ToString() + ")", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                Cap_Var_Form cvf = new Cap_Var_Form(Cells, lastSlaveReading, tech, test_type, cellLim, d_data, average, SWO1, slaveStepNum.ToString());
                                cvf.Owner = this;
                                cvf.Show();
                                cvf.Location = new Point(cvf.Location.X + 350, cvf.Location.Y + 120);

                            });
                        }

                    }
                    #endregion

                }
                catch (Exception ex)
                {
                    this.Invoke((MethodInvoker)delegate()
                    {
                        MessageBox.Show(this, "Error in master test poriton try:  " + ex.Message + Environment.NewLine + ex.TargetSite + Environment.NewLine + ex.HResult + Environment.NewLine + ex.StackTrace, "Main Test Catch Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }

            }, cRunTest[station].Token); // end thread

        }

        private void updateFinishTime(string MWO1, string MWO2, string MWO3, string SWO1, string SWO2, string SWO3, int stepNum, int stepNum2, int stepNum3, int slaveStepNum, int slaveStepNum2, int slaveStepNum3)
        {
            string strSETFINISHTIME = "UPDATE Tests SET DateCompleted= #" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "# WHERE WorkOrderNumber='" + MWO1.Trim() + "' AND StepNumber='" + stepNum.ToString("00") + "'";
            string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            OleDbConnection myAccessConn = new OleDbConnection(strAccessConn);

            OleDbCommand myAccessCommand2 = new OleDbCommand(strSETFINISHTIME, myAccessConn);

            lock (dataBaseLock)
            {
                myAccessConn.Open();
                myAccessCommand2.ExecuteNonQuery();
                myAccessConn.Close();
            }

            // do the same for MWO2 MWO3 and SW01 SW02 SW03

            if (MWO2 != "")
            {
                strSETFINISHTIME = "UPDATE Tests SET DateCompleted= #" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "# WHERE WorkOrderNumber='" + MWO2.Trim() + "' AND StepNumber='" + stepNum2.ToString("00") + "'";

                myAccessCommand2 = new OleDbCommand(strSETFINISHTIME, myAccessConn);

                lock (dataBaseLock)
                {
                    myAccessConn.Open();
                    myAccessCommand2.ExecuteNonQuery();
                    myAccessConn.Close();
                }

                if (MWO3 != "")
                {
                    strSETFINISHTIME = "UPDATE Tests SET DateCompleted= #" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "# WHERE WorkOrderNumber='" + MWO3.Trim() + "' AND StepNumber='" + stepNum3.ToString("00") + "'";

                    myAccessCommand2 = new OleDbCommand(strSETFINISHTIME, myAccessConn);

                    lock (dataBaseLock)
                    {
                        myAccessConn.Open();
                        myAccessCommand2.ExecuteNonQuery();
                        myAccessConn.Close();
                    }
                }
            }

            if (SWO1 != "")
            {
                strSETFINISHTIME = "UPDATE Tests SET DateCompleted= #" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "# WHERE WorkOrderNumber='" + SWO1.Trim() + "' AND StepNumber='" + slaveStepNum.ToString("00") + "'";

                myAccessCommand2 = new OleDbCommand(strSETFINISHTIME, myAccessConn);

                lock (dataBaseLock)
                {
                    myAccessConn.Open();
                    myAccessCommand2.ExecuteNonQuery();
                    myAccessConn.Close();
                }
                if (SWO2 != "")
                {
                    strSETFINISHTIME = "UPDATE Tests SET DateCompleted= #" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "# WHERE WorkOrderNumber='" + SWO2.Trim() + "' AND StepNumber='" + slaveStepNum2.ToString("00") + "'";

                    myAccessCommand2 = new OleDbCommand(strSETFINISHTIME, myAccessConn);

                    lock (dataBaseLock)
                    {
                        myAccessConn.Open();
                        myAccessCommand2.ExecuteNonQuery();
                        myAccessConn.Close();
                    }

                    if (SWO3 != "")
                    {
                        strSETFINISHTIME = "UPDATE Tests SET DateCompleted= #" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "# WHERE WorkOrderNumber='" + SWO3.Trim() + "' AND StepNumber='" + slaveStepNum3.ToString("00") + "'";

                        myAccessCommand2 = new OleDbCommand(strSETFINISHTIME, myAccessConn);

                        lock (dataBaseLock)
                        {
                            myAccessConn.Open();
                            myAccessCommand2.ExecuteNonQuery();
                            myAccessConn.Close();
                        }

                    }
                }

            }
        }// end RunTest

        private readonly object tLock = new object();
        bool testLock = false;

        private bool readTLock()
        {
            lock (tLock)
            {
                return testLock;
            }
        }

        private void setTLock()
        {
            lock (tLock)
            {
                testLock = true;
            }
        }

        private void clearTLock()
        {
            lock (tLock)
            {
                testLock = false;
            }
        }

        private void comboRunTest()
        {
            // we need to make sure that this is being run on an ICA with the auto config set to on
            // we got a combo test
            // set the station
            int station = dataGridView1.CurrentRow.Index;
            // we need to make sure that this is being run on an ICA with the auto config set to on
            if (!(GlobalVars.autoConfig && (bool)d.Rows[station][12] && (d.Rows[station][10].ToString().Contains("ICA") || d.Rows[station][10].ToString().Contains("MFC")) && !d.Rows[station][11].ToString().Contains("offline") && (bool)d.Rows[station][8]))
            {
                MessageBox.Show(this, "Cannot run a combo test, unless you are using a linked online ICA with AutoConfig", "Combo Test Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (d.Rows[station][6].ToString() == "")
            {
                sendNote(station, 3, "Combo Test Initiated");
            }
            else
            {
                sendNote(station, 3, "Combo Test Resumed");
            }

            // we are going to manage this on a helper thread...

            ThreadPool.QueueUserWorkItem(s =>
            {
                // SETUP //////////////////////////////////////////////////////////////////////////////////////////////////////
                //First look up the test info in the DB!!!!!
                // We need number of steps and all of the tests to run.....

                //first lets get the start temp and find if there is a master/slave relationship for special non-tests...
                //try to pull in the start temp
                double startTemp = (GlobalVars.CScanData[station].TP1 + GlobalVars.CScanData[station].TP2 + GlobalVars.CScanData[station].TP3 + GlobalVars.CScanData[station].TP4) / 4;
                #region master slave test check
                // we will use this bool to say if we need to do slave stuff..
                bool MasterSlaveTest = false;
                int slaveRow = -1;

                if (d.Rows[station][9].ToString().Length > 2)  // this is the case where we have a master and slave config
                {
                    MasterSlaveTest = true;
                    // also assign the slave channel...
                    string temp = d.Rows[station][9].ToString().Replace("-M", "");

                    for (int i = 0; i < 16; i++)
                    {
                        if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                        {
                            //found the slave
                            slaveRow = i;
                            break;
                        }
                    }
                }  // end the master slave find...

                #endregion

                // Now on to finding the specifics of the Combo Test
                OleDbConnection myAccessConn;
                string strAccessConn;

                try
                {
                    strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                    myAccessConn = new OleDbConnection(strAccessConn);
                }
                catch (Exception ex)
                {
                    this.Invoke((MethodInvoker)delegate()
                    {
                        MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                    cRunTest[station].Cancel();
                    return;
                }

                try
                {
                    // is this a resume or a new test?
                    string strAccessSelect;
                    int currentStep = 0;

                    if (d.Rows[station][2].ToString().Contains("("))
                    {
                        // we got a resume!
                        strAccessSelect = @"SELECT * FROM ComboTest WHERE ComboTestName='" + d.Rows[station][2].ToString().Substring(7, (d.Rows[station][2].ToString().IndexOf("(") - 8)) + "';";
                        currentStep = int.Parse(d.Rows[station][2].ToString().Substring(d.Rows[station][2].ToString().IndexOf("(") + 1, 2)) - 1;
                    }
                    else
                    {
                        // new test
                        strAccessSelect = @"SELECT * FROM ComboTest WHERE ComboTestName='" + d.Rows[station][2].ToString().Substring(7) + "';";
                        //currentStep = 0;
                    }


                    DataSet CustTest = new DataSet();
                    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                    lock (dataBaseLock)
                    {
                        myAccessConn.Open();
                        myDataAdapter.Fill(CustTest, "ComboTest");
                        myAccessConn.Close();
                    }


                    int steps = int.Parse(CustTest.Tables[0].Rows[0][2].ToString());

                    // here we go into the combo test loop...
                    for (int i = currentStep; i < steps; i++)
                    {
                        Debug.Print("step check...");
                        Debug.Print(currentStep.ToString());
                        Debug.Print(i.ToString());
                        Debug.Print(steps.ToString());

                        updateD(station, 2, "Combo: " + CustTest.Tables[0].Rows[0][1].ToString() + " (" + (i + 1).ToString() + " " + CustTest.Tables[0].Rows[0][i + 3].ToString() + ")");
                        if (MasterSlaveTest) { updateD(slaveRow, 2, "Combo: " + CustTest.Tables[0].Rows[0][1].ToString() + " (" + (i + 1).ToString() + " " + CustTest.Tables[0].Rows[0][i + 3].ToString() + ")"); }

                        // We have three cases.
                        // 1. Run a normal test
                        // 2. Do a wait
                        // 3. Do a temp settle

                        // 1. Run a normal test
                        if (CustTest.Tables[0].Rows[0][i + 3].ToString() != "Wait" && CustTest.Tables[0].Rows[0][i + 3].ToString() != "Temp Settle")
                        {
                            RunTest(CustTest.Tables[0].Rows[0][i + 3].ToString(), station);

                            // wait for the test to complete
                            // check for a cancel with the test marked as completed in the grid.. 
                            while (!cRunTest[station].IsCancellationRequested)
                            {
                                Thread.Sleep(100);
                            }
                            Thread.Sleep(100);

                            // Lets check if we should forceNext based on the test type.  This will allow us to do a shorting test after a Discharge fault
                            if (CustTest.Tables[0].Rows[0][i + 3].ToString().Contains("Discharge") && CustTest.Tables[0].Rows[0][i + 4].ToString().Contains("Shorting") && GlobalVars.advance2Short == true && d.Rows[station][11].ToString().Contains("Current Fault"))
                            {
                                updateD(station, 5, true);
                                if (MasterSlaveTest) { updateD(slaveRow, 5, true); }

                                forceNext = true;
                                // stop reset and clear also...
                                // first get the charger station
                                int Cstation = 0;
                                if (d.Rows[station][9].ToString() == "") { ;}  // do nothing if there is no assigned charger id (shouldn't happen with combo tests...)
                                else if (d.Rows[station][9].ToString().Length > 2)  // this is the case where we have a master and slave config
                                {

                                    MasterSlaveTest = true;
                                    // split into 3 and 4 digit case
                                    if (d.Rows[station][9].ToString().Length == 3)
                                    {
                                        // 3 case
                                        Cstation = int.Parse(d.Rows[station][9].ToString().Substring(0, 1));
                                    }
                                    else
                                    {
                                        // 4 case
                                        Cstation = int.Parse(d.Rows[station][9].ToString().Substring(0, 2));
                                    }

                                    // also assign the slave channel...
                                    string temp = d.Rows[station][9].ToString().Replace("-M", "");

                                    for (int ii = 0; ii < 16; ii++)
                                    {
                                        if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                                        {
                                            //found the slave
                                            slaveRow = i;
                                            break;
                                        }
                                    }
                                }  // end the master slave find...
                                else  // this is the normal case with just one charger
                                {
                                    Cstation = Convert.ToInt32(d.Rows[station][9]);
                                }

                                // now we need to reset the charger
                                updateD(station, 7, "Resetting Charger");
                                if (MasterSlaveTest) { updateD(slaveRow, 7, "Resetting Charger"); }

                                for (int j = 0; j < 10; j++)
                                {

                                    // set KE1 to 2 ("command")
                                    GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                                    // set KE3 to RESET
                                    GlobalVars.ICSettings[Cstation].KE3 = (byte)3;
                                    //Update the output string value
                                    GlobalVars.ICSettings[Cstation].UpdateOutText();
                                    criticalNum[Cstation] = true;
                                    //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                                    Thread.Sleep(5000);
                                    // set KE1 to 1 ("query")
                                    GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                                    //Update the output string value
                                    GlobalVars.ICSettings[Cstation].UpdateOutText();

                                    //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                                    for (int ii = 0; ii < 15; ii++)
                                    {
                                        criticalNum[Cstation] = true;
                                        Thread.Sleep(1000);
                                        if (d.Rows[station][11].ToString() == "RESET")
                                        {
                                            break;
                                        }
                                    }
                                    if (d.Rows[station][11].ToString() == "RESET")
                                    {
                                        break;
                                    }
                                }

                                // we are resuming after a fault has been corrected..
                                // now we need to clear the charger
                                updateD(station, 7, "Clearing Charger");
                                if (MasterSlaveTest) { updateD(slaveRow, 7, "Clearing Charger"); }


                                for (int j = 0; j < 10; j++)
                                {
                                    // set KE1 to 2 ("command")
                                    GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                                    // set KE3 to clear
                                    GlobalVars.ICSettings[Cstation].KE3 = (byte)0;
                                    //Update the output string value
                                    GlobalVars.ICSettings[Cstation].UpdateOutText();
                                    criticalNum[Cstation] = true;
                                    Thread.Sleep(3000);
                                    // set KE1 to 1 ("query")
                                    GlobalVars.ICSettings[Cstation].KE1 = (byte)0;

                                    //Update the output string value
                                    GlobalVars.ICSettings[Cstation].UpdateOutText();
                                    for (int ii = 0; ii < 3; ii++)
                                    {
                                        criticalNum[Cstation] = true;
                                        Thread.Sleep(1000);
                                        if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "RESET")
                                        {
                                            break;
                                        }
                                    }

                                    if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "RESET")
                                    {
                                        break;
                                    }
                                }

                                updateD(station, 5, false);
                                if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                            }



                            if (d.Rows[station][7].ToString() != "Complete" && forceNext == false)
                            {
                                // it was cancelled...
                                for (int j = 0; j < 200; j++)
                                {
                                    Thread.Sleep(100);
                                    if (d.Rows[station][7].ToString().Contains("Read ") || d.Rows[station][7].ToString().Contains("FAILED") || d.Rows[station][7].ToString() == "" || d.Rows[station][7].ToString().Contains("Test Cancelled"))
                                    {
                                        if (d.Rows[station][7].ToString() == "")
                                        {
                                            updateD(station, 2, "Combo: " + CustTest.Tables[0].Rows[0][1].ToString());
                                            if (MasterSlaveTest) { updateD(slaveRow, 2, "Combo: " + CustTest.Tables[0].Rows[0][1].ToString()); }
                                        }
                                        return; // return when the test is finished with clean up...
                                    }
                                }
                            }
                            else if (forceNext)
                            {
                                //wait for a Read or Test then go to next step...
                                forceNext = false;
                                Thread.Sleep(100);
                                while (!d.Rows[station][7].ToString().Contains("Test Cancelled") && !d.Rows[station][7].ToString().Contains("Read ") && d.Rows[station][7].ToString() != "Complete" && d.Rows[station][7].ToString() != "Clearing Charger")
                                {
                                    Thread.Sleep(100);
                                }
                                updateD(station, 5, true);
                                if (MasterSlaveTest) { updateD(slaveRow, 5, true); }
                                Thread.Sleep(2000);
                                updateD(station, 6, "");
                                if (MasterSlaveTest) { updateD(slaveRow, 6, ""); }
                            }

                        }
                        // 2. Do a wait
                        else if (CustTest.Tables[0].Rows[0][i + 3].ToString() == "Wait")
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                updateD(station, 5, true);
                                if (MasterSlaveTest) { updateD(slaveRow, 5, true); }
                                sendNote(station, 3, "Wait");
                                startNewTestToolStripMenuItem.Enabled = false;
                                resumeTestToolStripMenuItem.Enabled = false;
                                stopTestToolStripMenuItem.Enabled = true;
                                updateD(station, 7, "Wait");
                                if (MasterSlaveTest) { updateD(slaveRow, 7, "Wait"); }
                            });

                            cRunTest[station] = new CancellationTokenSource();
                            var stopwatch = new Stopwatch();

                            string eTimeS;



                            stopwatch.Start();

                            while (stopwatch.ElapsedMilliseconds < (int.Parse(CustTest.Tables[0].Rows[0][18].ToString()) * 60 * 1000))
                            {
                                Thread.Sleep(100);
                                if (cRunTest[station].IsCancellationRequested)
                                {
                                    if (forceNext)
                                    {

                                        updateD(station, 6, "");
                                        if (MasterSlaveTest) { updateD(slaveRow, 6, ""); }
                                        forceNext = false;
                                        updateD(station, 5, true);
                                        if (MasterSlaveTest) { updateD(slaveRow, 5, true); }
                                        Thread.Sleep(2000);
                                        break;
                                    }
                                    else
                                    {

                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            updateD(station, 5, false);
                                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                            updateD(station, 7, "Combo Cancelled");
                                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Combo Cancelled"); }
                                            sendNote(station, 2, "Combo Test Cancelled");
                                            startNewTestToolStripMenuItem.Enabled = true;
                                            resumeTestToolStripMenuItem.Enabled = true;
                                            stopTestToolStripMenuItem.Enabled = false;
                                        });
                                        return;
                                    }
                                }

                                eTimeS = stopwatch.Elapsed.ToString(@"hh\:mm\:ss");
                                updateD(station, 6, eTimeS);
                                if (MasterSlaveTest) { updateD(slaveRow, 6, eTimeS); }
                            }// end wait

                            updateD(station, 6, "");
                            if (MasterSlaveTest) { updateD(slaveRow, 6, ""); }
                            cRunTest[station].Cancel();
                        }
                        // 3. Do a temp settle
                        else
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                updateD(station, 5, true);
                                if (MasterSlaveTest) { updateD(slaveRow, 5, true); }
                                sendNote(station, 3, "Temp Settle");
                                startNewTestToolStripMenuItem.Enabled = false;
                                resumeTestToolStripMenuItem.Enabled = false;
                                stopTestToolStripMenuItem.Enabled = true;
                                updateD(station, 7, "Temp Settle");
                                if (MasterSlaveTest) { updateD(slaveRow, 7, "Temp Settle"); }
                            });

                            cRunTest[station] = new CancellationTokenSource();
                            var stopwatch = new Stopwatch();

                            string eTimeS;

                            updateD(station, 7, "Temp Settle");
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Temp Settle"); }

                            stopwatch.Start();

                            while (((GlobalVars.CScanData[station].TP1 + GlobalVars.CScanData[station].TP2 + GlobalVars.CScanData[station].TP3 + GlobalVars.CScanData[station].TP4) / 4) > (startTemp + int.Parse(CustTest.Tables[0].Rows[0][19].ToString())))
                            {
                                Thread.Sleep(100);
                                if (cRunTest[station].IsCancellationRequested)
                                {
                                    if (forceNext)
                                    {
                                        updateD(station, 6, "");
                                        if (MasterSlaveTest) { updateD(slaveRow, 6, ""); }
                                        forceNext = false;
                                        updateD(station, 5, true);
                                        if (MasterSlaveTest) { updateD(slaveRow, 5, true); }
                                        Thread.Sleep(2000);
                                        break;
                                    }
                                    else
                                    {

                                        this.Invoke((MethodInvoker)delegate()
                                        {
                                            updateD(station, 5, false);
                                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                                            updateD(station, 7, "Combo Cancelled");
                                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Combo Cancelled"); }
                                            sendNote(station, 2, "Combo Test Cancelled");
                                            startNewTestToolStripMenuItem.Enabled = true;
                                            resumeTestToolStripMenuItem.Enabled = true;
                                            stopTestToolStripMenuItem.Enabled = false;
                                        });
                                        return;
                                    }
                                }

                                eTimeS = stopwatch.Elapsed.ToString(@"hh\:mm\:ss");
                                updateD(station, 6, eTimeS);
                                if (MasterSlaveTest) { updateD(slaveRow, 6, eTimeS); }
                            }// end temp settle
                            cRunTest[station].Cancel();
                            updateD(station, 6, "");
                            if (MasterSlaveTest) { updateD(slaveRow, 6, ""); }

                        }


                    }// end for

                    // CLOSE UP SHOP ////////////////////////////////////////////////////////////////////////////////////////////
                    updateD(station, 2, "Combo: " + CustTest.Tables[0].Rows[0][1].ToString());
                    if (MasterSlaveTest) { updateD(slaveRow, 2, "Combo: " + CustTest.Tables[0].Rows[0][1].ToString()); }

                    this.Invoke((MethodInvoker)delegate()
                    {
                        sendNote(station, 2, "Combo Test Completed");
                    });

                    updateD(station, 5, false);
                    if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                }

                catch (Exception ex)
                {
                    this.Invoke((MethodInvoker)delegate()
                    {
                        MessageBox.Show(this, "Error: Something went wrong in the Combo Test Code. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                    updateD(station, 5, false);
                    if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                    return;
                }

            }); // end thread


        }

    }
}