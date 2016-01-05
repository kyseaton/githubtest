using System;
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

        private void RunTest()
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


            int station = dataGridView1.CurrentRow.Index;
            int Cstation = 0;

            //here to prevent double starts...
            if (cRunTest[station] != null && cRunTest[station].IsCancellationRequested == false)
            {
                return;
            }

            // gettting rid of isASlave
            //you can't run tests from a slave anymore..
            //bool isASlave = false;
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

            cRunTest[station] = new CancellationTokenSource();

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

            // Everything is going to be done on a helper thread
            ThreadPool.QueueUserWorkItem(s =>
            {
                string strAccessConn;
                string strAccessSelect;
                OleDbConnection myAccessConn;

                // setup the canellation token
                CancellationToken token = (CancellationToken)s;

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
                else if ((string)d.Rows[station][2] == "")
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
                else if (dataGridView1.Rows[station].Cells[4].Style.BackColor != Color.Green)
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
                else if (GlobalVars.CScanData[station].shuntCableType == "NONE" && (string)d.Rows[station][2] != "As Received")
                {
                    this.Invoke((MethodInvoker)delegate()
                    {
                        MessageBox.Show(this, "CScan is not connected to a cells Shunt cable.  Please connect a shunt cable to this CSCAN to run a test.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "Are you sure you want to continue? You have one work order assocaited with the master CScan, but have a 2X11 cable connected to it.  In order to record the data from all 22 channels you will need to add an additional workorder to this station.", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (dialogResult == DialogResult.No)
                    {
                        cRunTest[station].Cancel();
                        return;
                    }
                }
                else if ((MWO2 == "" || MWO3 == "") && GlobalVars.CScanData[station].CCID == 4)
                {
                    // warn the user that if they use a 3x7 cable with only one or two work orders that the data associated with the second or third battery will be lost
                    DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "Are you sure you want to continue? You have less than 3 work orders assocaited with the master CScan, but have a 3X7 cable connected to it.  In order to record the data from all 21 channels you will need to add an additional workorder(s) to this station.", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (dialogResult == DialogResult.No)
                    {
                        cRunTest[station].Cancel();
                        return;
                    }
                }


                //2x11 case
                if (GlobalVars.CScanData[station].CCID == 3 && (int)pci.Rows[station][3] != (GlobalVars.CScanData[station].cellsToDisplay / 2) && (int)pci.Rows[station][3] != -1)
                {
                    // warn the user that the number of cells set in the database does not match the work order...
                    DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The number of cells the battery contains does not match the cells cable currently being used.  Do you want to continue?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (dialogResult == DialogResult.No)
                    {
                        cRunTest[station].Cancel();
                        return;
                    }
                }
                //3X7 case
                if (GlobalVars.CScanData[station].CCID == 3 && (int)pci.Rows[station][3] != (GlobalVars.CScanData[station].cellsToDisplay / 3) && (int)pci.Rows[station][3] != -1)
                {
                    // warn the user that the number of cells set in the database does not match the work order...
                    DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The number of cells the battery contains does not match the cells cable currently being used.  Do you want to continue?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (dialogResult == DialogResult.No)
                    {
                        cRunTest[station].Cancel();
                        return;
                    }
                }
                //general case
                else if (pci.Rows[station][1].ToString().Contains("NiCd") && (int)pci.Rows[station][3] != GlobalVars.CScanData[station].cellsToDisplay && (int)pci.Rows[station][3] != -1)
                {
                    // warn the user that the number of cells set in the database does not match the work order...
                    DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The number of cells the battery contains does not match the cells cable currently being used.  Do you want to continue?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (dialogResult == DialogResult.No)
                    {
                        cRunTest[station].Cancel();
                        return;
                    }
                }

                // may also need to check some of this for the slave
                else if (MasterSlaveTest)
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
                    else if (dataGridView1.Rows[slaveRow].Cells[4].Style.BackColor != Color.Green)
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
                        DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "Are you sure you want to continue? You have one work order assocaited with the slave CScan, but have a 2X11 cable connected to it.  In order to record the data from all 22 channels you will need to add an additional workorder to this station.", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (dialogResult == DialogResult.No)
                        {
                            cRunTest[station].Cancel();
                            return;
                        }
                    }
                    else if ((SWO2 == "" || SWO3 == "") && GlobalVars.CScanData[slaveRow].CCID == 4)
                    {
                        // warn the user that if they use a 3x7 cable with only one or two work orders that the data associated with the second or third battery will be lost
                        DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "Are you sure you want to continue? You have less than 3 work orders assocaited with the slave CScan, but have a 3X7 cable connected to it.  In order to record the data from all 21 channels you will need to add an additional workorder(s) to this station.", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (dialogResult == DialogResult.No)
                        {
                            cRunTest[station].Cancel();
                            return;
                        }
                    }

                    //2x11 case
                    if (GlobalVars.CScanData[slaveRow].CCID == 3 && (int)pci.Rows[slaveRow][3] != (GlobalVars.CScanData[slaveRow].cellsToDisplay / 2) && (int)pci.Rows[station][3] != -1)
                    {
                        // warn the user that the number of cells set in the database does not match the work order...
                        DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The number of cells the battery connected to the slave CSCAN contains does not match the cells cable currently being used.  Do you want to continue?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (dialogResult == DialogResult.No)
                        {
                            cRunTest[station].Cancel();
                            return;
                        }
                    }
                    //3X7 case
                    if (GlobalVars.CScanData[slaveRow].CCID == 3 && (int)pci.Rows[slaveRow][3] != (GlobalVars.CScanData[slaveRow].cellsToDisplay / 3) && (int)pci.Rows[station][3] != -1)
                    {
                        // warn the user that the number of cells set in the database does not match the work order...
                        DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The number of cells the battery connected to the slave CSCAN contains does not match the cells cable currently being used.  Do you want to continue?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (dialogResult == DialogResult.No)
                        {
                            cRunTest[station].Cancel();
                            return;
                        }
                    }
                    //general case
                    else if (pci.Rows[slaveRow][1].ToString().Contains("NiCd") && (int)pci.Rows[slaveRow][3] != GlobalVars.CScanData[slaveRow].cellsToDisplay && (int)pci.Rows[station][3] != -1)
                    {
                        // warn the user that the number of cells set in the database does not match the work order...
                        DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The number of cells the battery connected to the slave CSCAN contains does not match the cells cable currently being used.  Do you want to continue?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (dialogResult == DialogResult.No)
                        {
                            cRunTest[station].Cancel();
                            return;
                        }
                    }

                    // check that the batteries are the same model!
                    // need to create a db connection
                    try
                    {
                        strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
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
                            DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The Master and Slave battery models do not match. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                            if (dialogResult == DialogResult.No)
                            {
                                cRunTest[station].Cancel();
                                return;
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

                //I removed this because the charger present will determine how the test is run...
                // also need to check if an intelligent charger is connected for autoconfig
                //else if (GlobalVars.autoConfig == true && d.Rows[station][9].ToString().Contains("ICA") && (string)d.Rows[station][2] != "As Received")
                //{
                //    MessageBox.Show(this, "Auto Configure is turned on, but there is no intelligent charger detected.  Please connect an intelligent charger or turn Auto Configure off in the tools menu.");
                //     return;
                //}


                bool runAsShunt = false;
                
                if (d.Rows[station][10].ToString().Contains("Shunt"))
                {
                    runAsShunt = true;
                }
                else if ((bool)d.Rows[station][8] == false && (string)d.Rows[station][2] != "As Received")
                {
                    DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The charger is not linked. Do you want to proceed with the test in shunt mode?", "Click Yes to continue or No to cancel the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (dialogResult == DialogResult.No)
                    {
                        cRunTest[station].Cancel();
                        return;
                    }
                    runAsShunt = true;
                }


                if ((string)d.Rows[station][11] == "offline!" && (string)d.Rows[station][2] != "As Received" && d.Rows[station][10].ToString().Contains("ICA") && !runAsShunt)
                {
                    this.Invoke((MethodInvoker)delegate()
                    {
                        MessageBox.Show(this, "The Intelligent Charger is set to be offline.  Please set the charger to be online by pressing the following key sequence on the charger: FUNC, 1, 1 and ENTER.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                    cRunTest[station].Cancel();
                    return;
                }

                else if ((d.Rows[station][9].ToString() == "" || d.Rows[station][10].ToString() == "") && (string)d.Rows[station][2] != "As Received")
                {
                    // we don't have a charger linked. Do we still want to continue...
                    this.Invoke((MethodInvoker)delegate()
                    {
                        MessageBox.Show(this, "There is no charger ID entered.  You must enter a charger ID to run any test other than 'As Received'", "Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    });
                    cRunTest[station].Cancel();
                    return;
                }


                //Finally Lets check if the charger is running also...
                else if (d.Rows[station][11].ToString() == "RUN" && d.Rows[station][10].ToString().Contains("ICA") && !runAsShunt)
                {
                    //looks like that charger is already running.  Lets ask the user if we should stop the charger or not.
                    DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The charger appears to already be running. Do you want to stop it now and proceed with the test?", "Click Yes to have the program stop the charger or No to do it manually", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (dialogResult == DialogResult.Yes)
                    {

                        // we'll check the box to indicate we are a go!
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
                    else
                    {
                        cRunTest[station].Cancel();
                        return;
                    }

                }
                #endregion

                #region db connection setup (all cases)
                // create db connection
                try
                {
                    strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
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
                if (d.Rows[station][10].ToString().Contains("ICA") && !runAsShunt)
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
                                    sendNote(station, 3, "Test Cancelled");
                                });

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

                #region if we are doing the autoconfig, let's get the charger settings in order and then loaded into the charger! (case 2 only)

                // we'll tell the charger what to do! (if we have an IC and the user wants us to...)
                if (GlobalVars.autoConfig && (bool)d.Rows[station][12] && d.Rows[station][10].ToString().Contains("ICA") && (string)d.Rows[station][2] != "As Received" && !runAsShunt)
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

                        // now we can assign the battery settings to the GlobalVars.
                        // We will decide on the settings based on the test being performed...
                        byte[] tempKMStore = new byte[21] { 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48, 48 }; // 21 values for the 21 KM params

                        #region setting switch
                        try
                        {
                            switch ((string)d.Rows[station][2])
                            {
                                case "Full Charge-6":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][42].ToString().Substring(0, 2)));          //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][43].ToString()));                          //time hours
                                    //tempKMStore[2] = (byte)(48);                                                                            //time mins
                                    if (d.Rows[station][10].ToString().Contains("mini"))
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][45].ToString()) * 10));             //top current byte
                                        tempKMStore[4] = (byte)(48);                                                                            //bottom current byte
                                    }
                                    else
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][45].ToString()) / 10));             //top current byte
                                        tempKMStore[4] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][45].ToString()) % 10));        //bottom current byte
                                    }
                                    tempKMStore[5] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][46].ToString()) / 1));              //top voltage byte
                                    tempKMStore[6] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][46].ToString()) % 1));            //bottom current byte

                                    tempKMStore[7] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][47].ToString()));                          //time 2 hours
                                    tempKMStore[8] = (byte)(48);                                                                            //time 2 mins
                                    if (d.Rows[station][10].ToString().Contains("mini"))
                                    {
                                        tempKMStore[9] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][49].ToString()) * 10));             //top current 2 byte
                                        tempKMStore[10] = (byte)(48);                                                                           //bottom current 2 byte
                                    }
                                    else
                                    {
                                        tempKMStore[9] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][49].ToString()) / 10));             //top current 2 byte
                                        tempKMStore[10] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][49].ToString()) % 10));       //bottom current 2 byte
                                    }
                                    tempKMStore[11] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][50].ToString()) / 1));                 //top voltage 2 byte
                                    tempKMStore[12] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][50].ToString()) % 1));           //bottom voltage 2 byte

                                    tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                    tempKMStore[14] = (byte)48;                                                                                //discharge time mins
                                    tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                    tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                    tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                    break;
                                case "Full Charge-4":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][52].ToString().Substring(0, 2)));          //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][53].ToString()));                          //time hours
                                    //tempKMStore[2] = (byte)(48);                                                                            //time mins
                                    if (d.Rows[station][10].ToString().Contains("mini"))
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][55].ToString()) * 10));             //top current byte
                                        tempKMStore[4] = (byte)(48);                                                                            //bottom current byte
                                    }
                                    else
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][55].ToString()) / 10));             //top current byte
                                        tempKMStore[4] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][55].ToString()) % 10));        //bottom current byte
                                    }
                                    tempKMStore[5] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][56].ToString()) / 1));                  //top voltage byte
                                    tempKMStore[6] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][56].ToString()) % 1));            //bottom current byte

                                    tempKMStore[7] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][57].ToString()));                          //time 2 hours
                                    tempKMStore[8] = (byte)(48);                                                                            //time 2 mins
                                    if (d.Rows[station][10].ToString().Contains("mini"))
                                    {
                                        tempKMStore[9] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][59].ToString()) * 10));             //top current 2 byte
                                        tempKMStore[10] = (byte)(48);                                                                           //bottom current 2 byte
                                    }
                                    else
                                    {
                                        tempKMStore[9] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][59].ToString()) / 10));             //top current 2 byte
                                        tempKMStore[10] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][59].ToString()) % 10));       //bottom current 2 byte
                                    }
                                    tempKMStore[11] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][60].ToString()) / 1));                 //top voltage 2 byte
                                    tempKMStore[12] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][60].ToString()) % 1));           //bottom voltage 2 byte

                                    tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                    tempKMStore[14] = (byte)48;                                                                              //discharge time mins
                                    tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                    tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                    tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                    break;
                                case "Top Charge-4":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][62].ToString().Substring(0, 2)));          //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][63].ToString()));                          //time hours
                                    tempKMStore[2] = (byte)(48);                                                                            //time mins
                                    if (d.Rows[station][10].ToString().Contains("mini"))
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][65].ToString()) * 10));             //top current byte
                                        tempKMStore[4] = (byte)(48);                                                                            //bottom current byte
                                    }
                                    else
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][65].ToString()) / 10));             //top current byte
                                        tempKMStore[4] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][65].ToString()) % 10));        //bottom current byte
                                    }
                                    tempKMStore[5] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][66].ToString()) / 1));              //top voltage byte
                                    tempKMStore[6] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][66].ToString()) % 1));            //bottom current byte

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
                                    break;
                                case "Top Charge-2":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][72].ToString().Substring(0, 2)));          //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][73].ToString()));                          //time hours
                                    tempKMStore[2] = (byte)(48);                                                                            //time mins
                                    if (d.Rows[station][10].ToString().Contains("mini"))
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][75].ToString()) * 10));             //top current byte
                                        tempKMStore[4] = (byte)(48);                                                                            //bottom current byte
                                    }
                                    else
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][75].ToString()) / 10));             //top current byte
                                        tempKMStore[4] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][75].ToString()) % 10));        //bottom current byte
                                    }
                                    tempKMStore[5] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][76].ToString()) / 1));                  //top voltage byte
                                    tempKMStore[6] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][76].ToString()) % 1));            //bottom current byte

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
                                    break;
                                case "Top Charge-1":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][82].ToString().Substring(0, 2)));          //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][83].ToString()));                          //time hours
                                    tempKMStore[2] = (byte)(48);                                                                            //time mins
                                    if (d.Rows[station][10].ToString().Contains("mini"))
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][85].ToString()) * 10));             //top current byte
                                        tempKMStore[4] = (byte)(48);                                                                            //bottom current byte
                                    }
                                    else
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][85].ToString()) / 10));             //top current byte
                                        tempKMStore[4] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][85].ToString()) % 10));        //bottom current byte
                                    }
                                    tempKMStore[5] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][86].ToString()) / 1));                  //top voltage byte
                                    tempKMStore[6] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][86].ToString()) % 1));            //bottom current byte

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
                                    break;
                                case "Capacity-1":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][92].ToString().Substring(0, 2)));          //mode
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

                                    tempKMStore[13] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][93].ToString()));                         //discharge time hours
                                    tempKMStore[14] = (byte)(48);                                                                           //discharge time mins
                                    if (tempKMStore[0] == 31 + 48)
                                    {
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[15] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][95].ToString()) * 1));         //discharge current high byte
                                            tempKMStore[16] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][95].ToString()) % 1));   //discharge current low byte
                                        }
                                        else
                                        {
                                            tempKMStore[15] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][95].ToString()) / 10));        //discharge current high byte
                                            tempKMStore[16] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][95].ToString()) % 10));   //discharge current low byte
                                        }
                                    }
                                    tempKMStore[17] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][96].ToString()) / 1));                 //discharge voltage high byte
                                    tempKMStore[18] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][96].ToString()) % 1));           //discharge voltage low byte
                                    if (tempKMStore[0] == 32 + 48)
                                    {
                                        tempKMStore[19] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][101].ToString()) / 1));            //discharge resistance high byte
                                        tempKMStore[20] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][101].ToString()) % 1));      //discharge resistance low byte
                                    }
                                    break;
                                case "Discharge":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][102].ToString().Substring(0, 2)));         //mode
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

                                    tempKMStore[13] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][103].ToString()));                        //discharge time hours
                                    tempKMStore[14] = (byte)(48);                                                                           //discharge time mins
                                    if (d.Rows[station][10].ToString().Contains("mini"))
                                    {
                                        tempKMStore[15] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][105].ToString()) * 1));            //discharge current high byte
                                        tempKMStore[16] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][105].ToString()) % 1));      //discharge current low byte
                                    }
                                    else
                                    {
                                        tempKMStore[15] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][105].ToString()) / 10));           //discharge current high byte
                                        tempKMStore[16] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][105].ToString()) % 10));      //discharge current low byte
                                    }
                                    tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                    break;
                                case "Slow Charge-14":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][112].ToString().Substring(0, 2)));         //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][113].ToString()));                         //time hours
                                    tempKMStore[2] = (byte)(48);                         //time mins
                                    if (d.Rows[station][10].ToString().Contains("mini"))
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][115].ToString()) * 10));            //top current byte
                                        tempKMStore[4] = (byte)(48);                                                                            //bottom current byte
                                    }
                                    else
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][115].ToString()) / 10));            //top current byte
                                        tempKMStore[4] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][115].ToString()) % 10));       //bottom current byte
                                    }
                                    tempKMStore[5] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][116].ToString()) / 1));                 //top voltage byte
                                    tempKMStore[6] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][116].ToString()) % 1));           //bottom current byte

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
                                    break;
                                case "Slow Charge-16":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][122].ToString().Substring(0, 2)));         //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][123].ToString()));                         //time hours
                                    tempKMStore[2] = (byte)(48);                         //time mins
                                    if (d.Rows[station][10].ToString().Contains("mini"))
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][125].ToString()) * 10));            //top current byte
                                        tempKMStore[4] = (byte)(48);                                                                            //bottom current byte
                                    }
                                    else
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][125].ToString()) / 10));            //top current byte
                                        tempKMStore[4] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][125].ToString()) % 10));       //bottom current byte
                                    }
                                    tempKMStore[5] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][126].ToString()) / 1));                 //top voltage byte
                                    tempKMStore[6] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][126].ToString()) % 1));           //bottom current byte

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
                                    break;
                                case "Custom Chg":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][132].ToString().Substring(0, 2)));         //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][133].ToString()));                         //time hours
                                    if (tempKMStore[0] != 20 && tempKMStore[0] != 21)
                                    {
                                        tempKMStore[2] = (byte)(48);                                                                        //time mins
                                    }
                                    if (d.Rows[station][10].ToString().Contains("mini"))
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][135].ToString()) * 10));            //top current byte
                                        tempKMStore[4] = (byte)(48);                                                                            //bottom current byte
                                    }
                                    else
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][135].ToString()) / 10));            //top current byte
                                        tempKMStore[4] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][135].ToString()) % 10));       //bottom current byte
                                    }
                                    tempKMStore[5] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][136].ToString()) / 1));                 //top voltage byte
                                    tempKMStore[6] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][136].ToString()) % 1));           //bottom current byte

                                    if (tempKMStore[0] == 20 || tempKMStore[0] == 21)
                                    {
                                        tempKMStore[7] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][137].ToString()));                     //time 2 hours
                                        tempKMStore[8] = (byte)(48);                                                                        //time 2 mins
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[9] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][139].ToString()) * 10));        //top current 2 byte
                                            tempKMStore[10] = (byte)(48);                                                                       //bottom current 2 byte
                                        }
                                        else
                                        {
                                            tempKMStore[9] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][139].ToString()) / 10));        //top current 2 byte
                                            tempKMStore[10] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][139].ToString()) % 10));  //bottom current 2 byte
                                        }
                                        tempKMStore[11] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][140].ToString()) / 1));            //top voltage 2 byte
                                        tempKMStore[12] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][140].ToString()) % 1));      //bottom voltage 2 byte
                                    }

                                    tempKMStore[13] = (byte)48;                                                                                //discharge time hours
                                    tempKMStore[14] = (byte)48;                                                                                //discharge time mins
                                    tempKMStore[15] = (byte)48;                                                                                //discharge current high byte
                                    tempKMStore[16] = (byte)48;                                                                                //discharge current low byte
                                    tempKMStore[17] = (byte)48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte)48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte)48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte)48;                                                                                //discharge resistance low byte
                                    break;
                                case "Custom Cap":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][142].ToString().Substring(0, 2)));         //mode
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

                                    tempKMStore[13] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][143].ToString()));                        //discharge time hours
                                    tempKMStore[14] = (byte)(48);                                                                           //discharge time mins
                                    if (tempKMStore[0] != 32 + 48)
                                    {
                                        if (d.Rows[station][10].ToString().Contains("mini"))
                                        {
                                            tempKMStore[15] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][145].ToString()) * 1));        //discharge current high byte
                                            tempKMStore[16] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][145].ToString()) % 1));  //discharge current low byte
                                        }
                                        else
                                        {
                                            tempKMStore[15] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][145].ToString()) / 10));       //discharge current high byte
                                            tempKMStore[16] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][145].ToString()) % 10));  //discharge current low byte
                                        }
                                    }
                                    if (tempKMStore[0] != 30 + 48)
                                    {
                                        tempKMStore[17] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][146].ToString()) / 1));            //discharge voltage high byte
                                        tempKMStore[18] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][146].ToString()) % 1));      //discharge voltage low byte
                                    }
                                    if (tempKMStore[0] == 32 + 48)
                                    {
                                        tempKMStore[19] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][151].ToString()) / 1));            //discharge resistance high byte
                                        tempKMStore[20] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][151].ToString()) % 1));      //discharge resistance low byte
                                    }
                                    break;
                                case "Constant Voltage":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][152].ToString().Substring(0, 2)));         //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][153].ToString()));                         //time hours
                                    tempKMStore[2] = (byte)(48);                                                                            //time mins
                                    if (d.Rows[station][10].ToString().Contains("mini"))
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][155].ToString()) * 10));            //top current byte
                                        tempKMStore[4] = (byte)(48);    //bottom current byte
                                    }
                                    else
                                    {
                                        tempKMStore[3] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][155].ToString()) / 10));            //top current byte
                                        tempKMStore[4] = (byte)(48 + 10 * (float.Parse(battery.Tables[0].Rows[0][155].ToString()) % 10));       //bottom current byte
                                    }
                                    tempKMStore[5] = (byte)(48 + (float.Parse(battery.Tables[0].Rows[0][156].ToString()) / 1));                 //top voltage byte
                                    tempKMStore[6] = (byte)(48 + 100 * (float.Parse(battery.Tables[0].Rows[0][156].ToString()) % 1));           //bottom voltage byte

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
                                    break;
                                default:
                                    break;
                            }// end switch
                        }
                        catch
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 3, "Error:  Fail to pull the settings from the DataBase.");
                                MessageBox.Show(this, "Fail to pull the settings from the DataBase. \r\nPlease make sure you have the battery model setup for this test under the Manage Battery Models menu.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });

                            //clear values from d
                            updateD(station, 7, ("Error"));
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                            clearTLock();
                            cRunTest[station].Cancel();
                            return;
                        }

                        #endregion

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

                    } // end try
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        this.Invoke((MethodInvoker)delegate()
                        {
                            sendNote(station, 3, "Error:  Failed to Auto Load settings into the Charger.");
                            MessageBox.Show(this, "Error: Failed to Auto Load settings into the Charger.\n", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        // reset everything
                        //clear values from d
                        updateD(station, 7, ("Error"));
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                        updateD(station, 5, false);
                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                        clearTLock();
                        cRunTest[station].Cancel();
                        return;
                    } // end catch

                }// end if

                #endregion

                #region load test readings and interval values (true for all cases)
                // Now we'll load the test parameters
                // We need to know the Interval and the number of readings///////////////////////////////////////////////////////////////////////

                int readings;
                int interval;

                //  open the db and pull in the options table
                try
                {
                    strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME='" + d.Rows[station][2].ToString() + "';";
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
                    this.Invoke((MethodInvoker)delegate()
                    {
                        sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                        MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                    updateD(station, 5, false);
                    if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                    //clear values from d
                    updateD(station, 7, ("Error"));
                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                    clearTLock();
                    cRunTest[station].Cancel();
                    return;
                }
                #endregion

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
                        this.Invoke((MethodInvoker)delegate()
                        {
                            sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        updateD(station, 5, false);
                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                        //clear values from d
                        updateD(station, 7, ("Error"));
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                        clearTLock();
                        cRunTest[station].Cancel();
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
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                            //clear values from d
                            updateD(station, 7, ("Error"));
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                            clearTLock();
                            cRunTest[station].Cancel();
                            return;
                        }
                    }




                    #endregion

                    // Save the test information to the test table///////////////////////////////////////////////////////////////////////////////////////////

                    // we need the technicial selected in the combo box too..
                    this.Invoke((MethodInvoker)delegate()
                    {
                        comboText = label2.Text;
                    });

                    #region save new test to test table
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
                            d.Rows[station][2].ToString() + "','" +                                //TestName
                            readings.ToString() + "','" +                                          //Reading
                            (interval * 1000).ToString() + "','" +                                 //interval in msec
                            station.ToString() + "','" +                                           // station number
                            d.Rows[station][10].ToString() + "',#" +                               // charger type
                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,#" +                 // start date
                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                 // date completed
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
                            d.Rows[station][2].ToString() + "','" +                                 //TestName
                            readings.ToString() + "','" +                                           //Reading
                            (interval * 1000).ToString() + "','" +                                  //interval in msec
                            station.ToString() + "','" +                                            // station number
                            d.Rows[station][10].ToString() + "',#" +                                // charger type
                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,#" +                  // start date
                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  // date completed
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
                            d.Rows[station][2].ToString() + "','" +                                 //TestName
                            readings.ToString() + "','" +                                           //Reading
                            (interval * 1000).ToString() + "','" +                                  //interval in msec
                            station.ToString() + "','" +                                            // station number
                            d.Rows[station][10].ToString() + "',#" +                                // charger type
                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,#" +                  // start date
                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  // date completed
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
                        this.Invoke((MethodInvoker)delegate()
                        {
                            sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        updateD(station, 5, false);
                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                        //clear values from d
                        updateD(station, 7, ("Error"));
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                        clearTLock();
                        cRunTest[station].Cancel();
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
                                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,#" +                 // start date
                                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                 // date completed
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
                                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,#" +                  // start date
                                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  // date completed
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
                                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,#" +                  // start date
                                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  // date completed
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
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                            //clear values from d
                            updateD(station, 7, ("Error"));
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                            clearTLock();
                            cRunTest[station].Cancel();
                            return;
                        }



                        //We made it to this point without errors, so we'll update the grid with the step number
                        updateD(slaveRow, 3, slaveStepNum.ToString());

                    }

                }// end if

                    #endregion

                else    // We've got a resume!
                {
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
                                stepNum2 = int.Parse((string)tests.Tables[0].Rows[0][4]) + 1;
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
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                            //clear values from d
                            updateD(station, 7, ("Error"));
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                            clearTLock();
                            cRunTest[station].Cancel();
                            return;
                        }
                    }


                    //do we need to do the same for the slave?
                    if (MasterSlaveTest)
                    {
                        slaveStepNum = int.Parse((string)d.Rows[slaveRow][3]);

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
                                slaveStepNum2 = int.Parse((string)tests.Tables[0].Rows[0][4]) + 1;
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
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                            //clear values from d
                            updateD(station, 7, ("Error"));
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                            clearTLock();
                            cRunTest[station].Cancel();
                            return;
                        }
                    }
                }

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
                    this.Invoke((MethodInvoker)delegate() { sendNote(station, 3, "Test Initiated"); });
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

                    if ((string)d.Rows[station][2] == "As Received")
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
                    return;
                }
                #endregion


                // OK now we'll tell the charger to startup (if we need to!)/////////////////////////////////////////////////////////////////////////////////////
                if (d.Rows[station][2].ToString() == "As Received")
                {
                    // nothing to do! if it's an "As Received" or we are running the test on a slave charger... 
                    stopwatch.Start();
                }
                else if (d.Rows[station][10].ToString().Contains("ICA") && !runAsShunt)
                {
                    #region             mode test
                    //update the GUI and pause...
                    updateD(station, 7, "Confirming Mode");
                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Confirming Mode"); }
                    Thread.Sleep(2000);

                    // we need to check that we are in the correct mode for the test...
                    string temp = GlobalVars.ICData[Cstation].testMode.ToString();
                    switch (d.Rows[station][2].ToString())
                    {
                        case "As Received":
                            // we should never get here..
                            break;
                        case "Full Charge-6":
                        case "Full Charge-4":
                            if (!(temp.Contains("20") || temp.Contains("21")))
                            {
                                DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The charger doesn't seem to be set up for a Full Charge. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dialogResult == DialogResult.No)
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
                                    return;
                                }  // end if

                            } // end if
                            break;
                        case "Top Charge-4":
                        case "Top Charge-2":
                        case "Top Charge-1":
                            if (!(temp.Contains("10") || temp.Contains("11")))
                            {
                                DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The charger doesn't seem to be set up for Top Charge. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dialogResult == DialogResult.No)
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
                                    return;
                                }  // end if
                            } // end if
                            break;
                        case "Constant Voltage":
                            if (!temp.Contains("12"))
                            {
                                DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The charger doesn't seem to be set up for Constant Voltage. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dialogResult == DialogResult.No)
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
                                    return;
                                }  // end if
                            } // end if
                            break;
                        case "Capacity-1":
                            if (!(temp.Contains("31") || temp.Contains("32")))
                            {
                                DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The charger doesn't seem to be set up for a Capacity test. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dialogResult == DialogResult.No)
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
                                    return;
                                }  // end if
                            } // end if
                            break;
                        case "Discharge":
                            if (!temp.Contains("30"))
                            {
                                DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The charger doesn't seem to be set up for Discharge. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dialogResult == DialogResult.No)
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
                                    return;
                                }  // end if
                            } // end if
                            break;
                        case "Slow Charge-14":
                        case "Slow Charge-16":
                            if (!(temp.Contains("10") || temp.Contains("11")))
                            {
                                DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The charger doesn't seem to be set up for Slow Charge. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dialogResult == DialogResult.No)
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
                                    return;
                                }  // end if
                            } // end if
                            break;
                        case "Custom Chg":
                            if (!(temp.Contains("10") || temp.Contains("11") || temp.Contains("12") || temp.Contains("20") || temp.Contains("21")))
                            {
                                DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The charger doesn't seem to be set up for a Custom Charge. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dialogResult == DialogResult.No)
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
                                    return;
                                }  // end if
                            } // end if
                            break;
                        case "Custom Cap":
                            if (!(temp.Contains("30") || temp.Contains("31") || temp.Contains("32")))
                            {
                                DialogResult dialogResult = MessageBox.Show(new Form() { TopMost = true }, "The charger doesn't seem to be set up for a Custom Capacity test. Are you sure you want to proceed with the test?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dialogResult == DialogResult.No)
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
                            sendNote(station, 3, "Test Cancelled");
                            startNewTestToolStripMenuItem.Enabled = true;
                            resumeTestToolStripMenuItem.Enabled = true;
                            stopTestToolStripMenuItem.Enabled = false;
                        });

                        clearTLock();
                        cRunTest[station].Cancel();
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
                            sendNote(station, 3, "Test Cancelled");
                            startNewTestToolStripMenuItem.Enabled = true;
                            resumeTestToolStripMenuItem.Enabled = true;
                            stopTestToolStripMenuItem.Enabled = false;
                        });
                        clearTLock();
                        cRunTest[station].Cancel();
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
                            // set KE3 to run
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
                                if (d.Rows[station][11].ToString() == "HOLD")
                                {
                                    break;
                                }
                            }

                            //make sure the charger has priority
                            if (d.Rows[station][11].ToString() == "HOLD")
                            {
                                break;
                            }
                        }

                    }

                    //cancel test
                    #region cancel block
                    if (token.IsCancellationRequested)
                    {

                        if ((string)d.Rows[station][2] == "As Received")
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
                        return;
                    }
                    #endregion

                    updateD(station, 7, "Telling Charger to Run");
                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Telling Charger to Run"); }
                    Thread.Sleep(2000);
                    // we are going to use a thread in a thread to start the charger up
                    // maybe this will reduce the timing difference!
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

                    // start timer now...
                    Thread.Sleep(1000);

                    //cancel test
                    #region cancel block
                    if (token.IsCancellationRequested)
                    {
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
                        return;
                    }
                    #endregion

                    stopwatch.Start();
                    Thread.Sleep(200);


                }// end else if for ICs...
                else if (d.Rows[station][10].ToString().Contains("CCA") && !runAsShunt)
                {
                    // We have a legacy Charger!
                    // We need to let it run!
                    
                    //cancel test
                    #region cancel block
                    if (token.IsCancellationRequested)
                    {

                        if ((string)d.Rows[station][2] == "As Received")
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

                            //update the gui
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 3, "Test Cancelled");
                                startNewTestToolStripMenuItem.Enabled = true;
                                resumeTestToolStripMenuItem.Enabled = true;
                                stopTestToolStripMenuItem.Enabled = false;
                            });

                            clearTLock();
                            return;
                        }
                        #endregion
                        // look for a current
                        if (Math.Abs(GlobalVars.CScanData[station].currentOne) > 0.2)
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
                    this.Invoke((MethodInvoker)delegate()
                    {
                        MessageBox.Show(this, "No Charger Detected!  Please check settings!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });

                    //clear values from d
                    updateD(station, 7, ("No Charger Detected!"));
                    if (MasterSlaveTest) { updateD(slaveRow, 7, "No Charger Detected!"); }
                    updateD(station, 5, false);
                    if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                    //update the gui
                    this.Invoke((MethodInvoker)delegate()
                    {
                        sendNote(station, 3, "No Charger Detected");
                        startNewTestToolStripMenuItem.Enabled = true;
                        resumeTestToolStripMenuItem.Enabled = true;
                        stopTestToolStripMenuItem.Enabled = false;
                    });

                    clearTLock();   // for good luck!
                    cRunTest[station].Cancel();
                    return;
                }

                // we made it.
                //tell them what reading we are on!
                updateD(station, 7, ("Reading " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                if (MasterSlaveTest) { updateD(slaveRow, 7, ("Reading " + (currentReading - 1).ToString() + " of " + readings.ToString())); }

                bool startUpDelay = true;
                // clear startUpDelay in 15 seconds...
                ThreadPool.QueueUserWorkItem(t =>
                {
                    Thread.Sleep(10000);
                    startUpDelay = false;
                });


                clearTLock();

                int badCurCount = 0;


                while (currentReading <= readings)
                {
                    // check if we need to take a reading
                    if (((currentReading - 1) * interval * 1000) < stopwatch.Elapsed.Add(offset).TotalMilliseconds)
                    {
                        //first record the elapsed amount of time
                        TimeSpan temp = stopwatch.Elapsed.Add(offset);
                        // update the grid
                        updateD(station, 7, ("Reading " + currentReading.ToString() + " of " + readings.ToString()));
                        if (MasterSlaveTest) { updateD(slaveRow, 7, ("Reading " + currentReading.ToString() + " of " + readings.ToString())); }

                        #region save a scan to the DB
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
                                GlobalVars.CScanData[station].orderedCells[11].ToString("0.000") + "','" +      //CEL01
                                GlobalVars.CScanData[station].orderedCells[12].ToString("0.000") + "','" +      //CEL02
                                GlobalVars.CScanData[station].orderedCells[13].ToString("0.000") + "','" +      //CEL03
                                GlobalVars.CScanData[station].orderedCells[14].ToString("0.000") + "','" +      //CEL04
                                GlobalVars.CScanData[station].orderedCells[15].ToString("0.000") + "','" +      //CEL05
                                GlobalVars.CScanData[station].orderedCells[16].ToString("0.000") + "','" +      //CEL06
                                GlobalVars.CScanData[station].orderedCells[17].ToString("0.000") + "','" +      //CEL07
                                GlobalVars.CScanData[station].orderedCells[18].ToString("0.000") + "','" +      //CEL08
                                GlobalVars.CScanData[station].orderedCells[19].ToString("0.000") + "','" +      //CEL09
                                GlobalVars.CScanData[station].orderedCells[20].ToString("0.000") + "','" +      //CEL10
                                GlobalVars.CScanData[station].orderedCells[21].ToString("0.000") + "','" +     //CEL11
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
                                GlobalVars.CScanData[station].orderedCells[7].ToString("0.000") + "','" +      //CEL01
                                GlobalVars.CScanData[station].orderedCells[8].ToString("0.000") + "','" +      //CEL02
                                GlobalVars.CScanData[station].orderedCells[9].ToString("0.000") + "','" +      //CEL03
                                GlobalVars.CScanData[station].orderedCells[10].ToString("0.000") + "','" +      //CEL04
                                GlobalVars.CScanData[station].orderedCells[11].ToString("0.000") + "','" +      //CEL05
                                GlobalVars.CScanData[station].orderedCells[12].ToString("0.000") + "','" +      //CEL06
                                GlobalVars.CScanData[station].orderedCells[13].ToString("0.000") + "','" +      //CEL07
                                GlobalVars.CScanData[station].orderedCells[14].ToString("0.000") + "','" +      //CEL08
                                GlobalVars.CScanData[station].orderedCells[15].ToString("0.000") + "','" +      //CEL09
                                GlobalVars.CScanData[station].orderedCells[16].ToString("0.000") + "','" +      //CEL10
                                GlobalVars.CScanData[station].orderedCells[17].ToString("0.000") + "','" +     //CEL11
                                GlobalVars.CScanData[station].orderedCells[18].ToString("0.000") + "','" +     //CEL12
                                GlobalVars.CScanData[station].orderedCells[19].ToString("0.000") + "','" +     //CEL13
                                GlobalVars.CScanData[station].orderedCells[20].ToString("0.000") + "','" +     //CEL14
                                GlobalVars.CScanData[station].orderedCells[21].ToString("0.000") + "','" +     //CEL15
                                GlobalVars.CScanData[station].orderedCells[22].ToString("0.000") + "','" +     //CEL16
                                GlobalVars.CScanData[station].orderedCells[23].ToString("0.000") + "','" +     //CEL17
                                GlobalVars.CScanData[station].orderedCells[0].ToString("0.000") + "','" +     //CEL18
                                GlobalVars.CScanData[station].orderedCells[1].ToString("0.000") + "','" +     //CEL19
                                GlobalVars.CScanData[station].orderedCells[2].ToString("0.000") + "','" +     //CEL20
                                GlobalVars.CScanData[station].orderedCells[3].ToString("0.000") + "','" +     //CEL21
                                GlobalVars.CScanData[station].orderedCells[4].ToString("0.000") + "','" +     //CEL22
                                GlobalVars.CScanData[station].orderedCells[5].ToString("0.000") + "','" +     //CEL23
                                GlobalVars.CScanData[station].orderedCells[6].ToString("0.000") + "','" +     //CEL24
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
                                GlobalVars.CScanData[station].orderedCells[14].ToString("0.000") + "','" +      //CEL01
                                GlobalVars.CScanData[station].orderedCells[15].ToString("0.000") + "','" +      //CEL02
                                GlobalVars.CScanData[station].orderedCells[16].ToString("0.000") + "','" +      //CEL03
                                GlobalVars.CScanData[station].orderedCells[17].ToString("0.000") + "','" +      //CEL04
                                GlobalVars.CScanData[station].orderedCells[18].ToString("0.000") + "','" +      //CEL05
                                GlobalVars.CScanData[station].orderedCells[19].ToString("0.000") + "','" +      //CEL06
                                GlobalVars.CScanData[station].orderedCells[20].ToString("0.000") + "','" +      //CEL07
                                GlobalVars.CScanData[station].orderedCells[21].ToString("0.000") + "','" +      //CEL08
                                GlobalVars.CScanData[station].orderedCells[22].ToString("0.000") + "','" +      //CEL09
                                GlobalVars.CScanData[station].orderedCells[23].ToString("0.000") + "','" +      //CEL10
                                GlobalVars.CScanData[station].orderedCells[0].ToString("0.000") + "','" +     //CEL11
                                GlobalVars.CScanData[station].orderedCells[1].ToString("0.000") + "','" +     //CEL12
                                GlobalVars.CScanData[station].orderedCells[2].ToString("0.000") + "','" +     //CEL13
                                GlobalVars.CScanData[station].orderedCells[3].ToString("0.000") + "','" +     //CEL14
                                GlobalVars.CScanData[station].orderedCells[4].ToString("0.000") + "','" +     //CEL15
                                GlobalVars.CScanData[station].orderedCells[5].ToString("0.000") + "','" +     //CEL16
                                GlobalVars.CScanData[station].orderedCells[6].ToString("0.000") + "','" +     //CEL17
                                GlobalVars.CScanData[station].orderedCells[7].ToString("0.000") + "','" +     //CEL18
                                GlobalVars.CScanData[station].orderedCells[8].ToString("0.000") + "','" +     //CEL19
                                GlobalVars.CScanData[station].orderedCells[9].ToString("0.000") + "','" +     //CEL20
                                GlobalVars.CScanData[station].orderedCells[10].ToString("0.000") + "','" +     //CEL21
                                GlobalVars.CScanData[station].orderedCells[11].ToString("0.000") + "','" +     //CEL22
                                GlobalVars.CScanData[station].orderedCells[12].ToString("0.000") + "','" +     //CEL23
                                GlobalVars.CScanData[station].orderedCells[13].ToString("0.000") + "','" +     //CEL24
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

                            //also insert the slave reading is need be...
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
                                        GlobalVars.CScanData[slaveRow].orderedCells[11].ToString("0.000") + "','" +      //CEL01
                                        GlobalVars.CScanData[slaveRow].orderedCells[12].ToString("0.000") + "','" +      //CEL02
                                        GlobalVars.CScanData[slaveRow].orderedCells[13].ToString("0.000") + "','" +      //CEL03
                                        GlobalVars.CScanData[slaveRow].orderedCells[14].ToString("0.000") + "','" +      //CEL04
                                        GlobalVars.CScanData[slaveRow].orderedCells[15].ToString("0.000") + "','" +      //CEL05
                                        GlobalVars.CScanData[slaveRow].orderedCells[16].ToString("0.000") + "','" +      //CEL06
                                        GlobalVars.CScanData[slaveRow].orderedCells[17].ToString("0.000") + "','" +      //CEL07
                                        GlobalVars.CScanData[slaveRow].orderedCells[18].ToString("0.000") + "','" +      //CEL08
                                        GlobalVars.CScanData[slaveRow].orderedCells[19].ToString("0.000") + "','" +      //CEL09
                                        GlobalVars.CScanData[slaveRow].orderedCells[20].ToString("0.000") + "','" +      //CEL10
                                        GlobalVars.CScanData[slaveRow].orderedCells[21].ToString("0.000") + "','" +     //CEL11
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
                                        GlobalVars.CScanData[slaveRow].orderedCells[22].ToString("0.000") + "','" +     //CEL23
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
                                    GlobalVars.CScanData[slaveRow].orderedCells[7].ToString("0.000") + "','" +      //CEL01
                                    GlobalVars.CScanData[slaveRow].orderedCells[8].ToString("0.000") + "','" +      //CEL02
                                    GlobalVars.CScanData[slaveRow].orderedCells[9].ToString("0.000") + "','" +      //CEL03
                                    GlobalVars.CScanData[slaveRow].orderedCells[10].ToString("0.000") + "','" +      //CEL04
                                    GlobalVars.CScanData[slaveRow].orderedCells[11].ToString("0.000") + "','" +      //CEL05
                                    GlobalVars.CScanData[slaveRow].orderedCells[12].ToString("0.000") + "','" +      //CEL06
                                    GlobalVars.CScanData[slaveRow].orderedCells[13].ToString("0.000") + "','" +      //CEL07
                                    GlobalVars.CScanData[slaveRow].orderedCells[14].ToString("0.000") + "','" +      //CEL08
                                    GlobalVars.CScanData[slaveRow].orderedCells[15].ToString("0.000") + "','" +      //CEL09
                                    GlobalVars.CScanData[slaveRow].orderedCells[16].ToString("0.000") + "','" +      //CEL10
                                    GlobalVars.CScanData[slaveRow].orderedCells[17].ToString("0.000") + "','" +     //CEL11
                                    GlobalVars.CScanData[slaveRow].orderedCells[18].ToString("0.000") + "','" +     //CEL12
                                    GlobalVars.CScanData[slaveRow].orderedCells[19].ToString("0.000") + "','" +     //CEL13
                                    GlobalVars.CScanData[slaveRow].orderedCells[20].ToString("0.000") + "','" +     //CEL14
                                    GlobalVars.CScanData[slaveRow].orderedCells[21].ToString("0.000") + "','" +     //CEL15
                                    GlobalVars.CScanData[slaveRow].orderedCells[22].ToString("0.000") + "','" +     //CEL16
                                    GlobalVars.CScanData[slaveRow].orderedCells[23].ToString("0.000") + "','" +     //CEL17
                                    GlobalVars.CScanData[slaveRow].orderedCells[0].ToString("0.000") + "','" +     //CEL18
                                    GlobalVars.CScanData[slaveRow].orderedCells[1].ToString("0.000") + "','" +     //CEL19
                                    GlobalVars.CScanData[slaveRow].orderedCells[2].ToString("0.000") + "','" +     //CEL20
                                    GlobalVars.CScanData[slaveRow].orderedCells[3].ToString("0.000") + "','" +     //CEL21
                                    GlobalVars.CScanData[slaveRow].orderedCells[4].ToString("0.000") + "','" +     //CEL22
                                    GlobalVars.CScanData[slaveRow].orderedCells[5].ToString("0.000") + "','" +     //CEL23
                                    GlobalVars.CScanData[slaveRow].orderedCells[6].ToString("0.000") + "','" +     //CEL24
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
                                    GlobalVars.CScanData[slaveRow].orderedCells[14].ToString("0.000") + "','" +      //CEL01
                                    GlobalVars.CScanData[slaveRow].orderedCells[15].ToString("0.000") + "','" +      //CEL02
                                    GlobalVars.CScanData[slaveRow].orderedCells[16].ToString("0.000") + "','" +      //CEL03
                                    GlobalVars.CScanData[slaveRow].orderedCells[17].ToString("0.000") + "','" +      //CEL04
                                    GlobalVars.CScanData[slaveRow].orderedCells[18].ToString("0.000") + "','" +      //CEL05
                                    GlobalVars.CScanData[slaveRow].orderedCells[19].ToString("0.000") + "','" +      //CEL06
                                    GlobalVars.CScanData[slaveRow].orderedCells[20].ToString("0.000") + "','" +      //CEL07
                                    GlobalVars.CScanData[slaveRow].orderedCells[21].ToString("0.000") + "','" +      //CEL08
                                    GlobalVars.CScanData[slaveRow].orderedCells[22].ToString("0.000") + "','" +      //CEL09
                                    GlobalVars.CScanData[slaveRow].orderedCells[23].ToString("0.000") + "','" +      //CEL10
                                    GlobalVars.CScanData[slaveRow].orderedCells[0].ToString("0.000") + "','" +     //CEL11
                                    GlobalVars.CScanData[slaveRow].orderedCells[1].ToString("0.000") + "','" +     //CEL12
                                    GlobalVars.CScanData[slaveRow].orderedCells[2].ToString("0.000") + "','" +     //CEL13
                                    GlobalVars.CScanData[slaveRow].orderedCells[3].ToString("0.000") + "','" +     //CEL14
                                    GlobalVars.CScanData[slaveRow].orderedCells[4].ToString("0.000") + "','" +     //CEL15
                                    GlobalVars.CScanData[slaveRow].orderedCells[5].ToString("0.000") + "','" +     //CEL16
                                    GlobalVars.CScanData[slaveRow].orderedCells[6].ToString("0.000") + "','" +     //CEL17
                                    GlobalVars.CScanData[slaveRow].orderedCells[7].ToString("0.000") + "','" +     //CEL18
                                    GlobalVars.CScanData[slaveRow].orderedCells[8].ToString("0.000") + "','" +     //CEL19
                                    GlobalVars.CScanData[slaveRow].orderedCells[9].ToString("0.000") + "','" +     //CEL20
                                    GlobalVars.CScanData[slaveRow].orderedCells[10].ToString("0.000") + "','" +     //CEL21
                                    GlobalVars.CScanData[slaveRow].orderedCells[11].ToString("0.000") + "','" +     //CEL22
                                    GlobalVars.CScanData[slaveRow].orderedCells[12].ToString("0.000") + "','" +     //CEL23
                                    GlobalVars.CScanData[slaveRow].orderedCells[13].ToString("0.000") + "','" +     //CEL24
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
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 3, "Error: Failed to retrieve the required data from the DataBase.");
                                startNewTestToolStripMenuItem.Enabled = true;
                                resumeTestToolStripMenuItem.Enabled = true;
                                stopTestToolStripMenuItem.Enabled = false;
                                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                            //clear values from d
                            updateD(station, 7, ("Error"));
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Error"); }
                            cRunTest[station].Cancel();
                            return;
                        }
                        #endregion

                        // finally update the reading
                        currentReading++;
                    }

                    //Now update the timer
                    eTime = stopwatch.Elapsed.Add(offset);
                    eTimeS = eTime.ToString(@"hh\:mm\:ss");
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
                    if (d.Rows[station][10].ToString().Contains("ICA") && !startUpDelay && !runAsShunt)
                    {
                        //current check part...
                        //if we have a mini that is charging...
                        if (d.Rows[station][10].ToString().Contains("mini") && !(d.Rows[station][2].ToString().Contains("Cap") || d.Rows[station][2].ToString().Contains("Discharge")))
                        {
                            //badCurCount looks at current two...
                            // for the mini case
                            if (Math.Abs(GlobalVars.CScanData[station].currentTwo) < 0.05)
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


                        if (badCurCount > 100)
                        {
                            // end the test!
                            //clear values from d

                            //update the gui
                            this.Invoke((MethodInvoker)delegate()
                            {
                                sendNote(station, 3, "Test failed. Charger is not producing any current.  Please check the charger settings.");
                                startNewTestToolStripMenuItem.Enabled = true;
                                resumeTestToolStripMenuItem.Enabled = true;
                                stopTestToolStripMenuItem.Enabled = false;
                                MessageBox.Show(this, "Test failed. Charger is not producing any current.  Please check the charger settings.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });

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
                                    if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END")
                                    {
                                        break;
                                    }

                                }

                                if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END")
                                {
                                    break;
                                }

                            }

                            updateD(station, 7, ("FAILED ON " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                            if (MasterSlaveTest) { updateD(slaveRow, 7, ("FAILED ON " + (currentReading - 1).ToString() + " of " + readings.ToString())); }
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                            cRunTest[station].Cancel();
                            return;
                        }


                        ////////status checks!
                        
                        if ((currentReading == readings || currentReading == readings + 1) && d.Rows[station][11].ToString() == "END")
                        {

                            // we are on the last reading.  Look out for the "END"...
                            // At the moment the test will just continue until the time is up...

                        }
                        else if ((string)d.Rows[station][11] != "RUN" && (string)d.Rows[station][2] != "As Received")
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

                                if ((string)d.Rows[station][11] == "Power Fail" || (string)d.Rows[station][11] == "HOLD")
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

                                            //update the gui
                                            this.Invoke((MethodInvoker)delegate()
                                            {
                                                sendNote(station, 3, "Test Cancelled");
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
                                    // end the test!
                                    //clear values from d

                                    updateD(station, 7, ("FAILED ON " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                                    if (MasterSlaveTest) { updateD(slaveRow, 7, ("FAILED ON " + (currentReading - 1).ToString() + " of " + readings.ToString())); }
                                    updateD(station, 5, false);
                                    if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

                                    //update the gui
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        sendNote(station, 3, "Test failed. Charger Status:  " + d.Rows[station][11].ToString());
                                        startNewTestToolStripMenuItem.Enabled = true;
                                        resumeTestToolStripMenuItem.Enabled = true;
                                        stopTestToolStripMenuItem.Enabled = false;
                                        MessageBox.Show(this, "Test failed. Charger Status:  " + d.Rows[station][11].ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    });
                                    cRunTest[station].Cancel();
                                    return;
                                }// end test fail else

                            }  // end still no run if
                        } // end no run if
                    }// end IC block
                    else if (d.Rows[station][10].ToString().Contains("CCA") && !startUpDelay && !runAsShunt)
                    {
                        // check that the C-Scan is still running...

                        // this is the power fail case...
                        if (d.Rows[station][11].ToString() == "Power Fail")
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
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        sendNote(station, 3, "Test canceled while in a power fail.");
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
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        sendNote(station, 3, "Current has fallen below minimum threshold. Please check the shunt connection and resume or restart the test.");
                                        startNewTestToolStripMenuItem.Enabled = true;
                                        resumeTestToolStripMenuItem.Enabled = true;
                                        stopTestToolStripMenuItem.Enabled = false;
                                        MessageBox.Show(this, "Current has fallen below minimum threshold. Please check the shunt connection and resume or restart the test.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    });
                                    cRunTest[station].Cancel();
                                    return;
                                }
                            }
                        }


                    }
                    else if ((d.Rows[station][10].ToString().Contains("Shunt") || runAsShunt) && !startUpDelay)
                    {
                        //test that the current is still above the 0.2 threshold...
                        // check that the C-Scan is still running...
                        if (Math.Abs(GlobalVars.CScanData[station].currentOne) < 0.2)
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

                                    //update the gui
                                    this.Invoke((MethodInvoker)delegate()
                                    {
                                        sendNote(station, 3, "Current has fallen below minimum threshold.  Please check the shunt connection and resume or restart the test.");
                                        startNewTestToolStripMenuItem.Enabled = true;
                                        resumeTestToolStripMenuItem.Enabled = true;
                                        stopTestToolStripMenuItem.Enabled = false;
                                        MessageBox.Show(this, "Current has fallen below minimum threshold. Please check the shunt connection and resume or restart the test.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    });
                                    cRunTest[station].Cancel();
                                    return;
                                }
                            }
                        }

                    }
                    else if (d.Rows[station][2].ToString().Contains("As Received"))
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
                        this.Invoke((MethodInvoker)delegate()
                        {
                            sendNote(station, 3, "Charger disconnected.  Please check connection and resume or restart the test.");
                            startNewTestToolStripMenuItem.Enabled = true;
                            resumeTestToolStripMenuItem.Enabled = true;
                            stopTestToolStripMenuItem.Enabled = false;
                            MessageBox.Show(this, "Charger disconnected.  Please check connection and resume or restart the test.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        cRunTest[station].Cancel();
                        return;

                    }


                    #endregion

                    //Now we should check for a cancel
                    #region cancel block
                    if (token.IsCancellationRequested)
                    {

                        if ((string)d.Rows[station][2] == "As Received")
                        {
                            //nothing to do here...
                        }
                        else if (d.Rows[station][10].ToString().Contains("ICA") && !runAsShunt)
                        {
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

                                for (int i = 0; i < 15; i++)
                                {
                                    criticalNum[Cstation] = true;
                                    Thread.Sleep(1000);
                                    if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END")
                                    {
                                        break;
                                    }

                                }

                                if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END")
                                {
                                    break;
                                }

                            }

                        }
                        else if (d.Rows[station][10].ToString().Contains("CCA") && !runAsShunt)
                        {
                            // Put the charger back on hold...
                            GlobalVars.cHold[station] = true;
                        }
                        else if (d.Rows[station][10].ToString().Contains("Shunt") ||  runAsShunt)
                        {
                            // Also nothing to do...
                        }

                        //clear values from d
                        updateD(station, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                        if (MasterSlaveTest) { updateD(slaveRow, 7, ("Read " + (currentReading - 1).ToString() + " of " + readings.ToString())); }
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

                        return;
                    }
                    #endregion

                    //every interval is defined in seconds to be safe, we'll test if we are at the correct interval every 200ms
                    Thread.Sleep(200);

                } // end main test loop...


                // We finished so let's clearn up!
                // If we are running the charger tell it to stop and reset
                if ((string)d.Rows[station][2] == "As Received")
                {
                    // nothing to do...
                }
                else if ((string)d.Rows[station][9] != "" && d.Rows[station][10].ToString().Contains("ICA") && !runAsShunt)
                {
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

                        for (int i = 0; i < 15; i++)
                        {
                            criticalNum[Cstation] = true;
                            Thread.Sleep(1000);
                            if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END")
                            {
                                break;
                            }

                        }

                        if (d.Rows[station][11].ToString() == "HOLD" || d.Rows[station][11].ToString() == "END")
                        {
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
                else if (d.Rows[station][10].ToString().Contains("Shunt") ||  runAsShunt)
                {
                    // Also nothing to do...
                }


                //update the gui
                this.Invoke((MethodInvoker)delegate()
                {
                    sendNote(station, 3, "Test Complete");
                    startNewTestToolStripMenuItem.Enabled = true;
                    resumeTestToolStripMenuItem.Enabled = false;
                    stopTestToolStripMenuItem.Enabled = false;
                });

                //Test is finished!
                updateD(station, 6, "");
                if (MasterSlaveTest) { updateD(slaveRow, 6, ""); }
                updateD(station, 7, "Complete");
                if (MasterSlaveTest) { updateD(slaveRow, 7, "Complete"); }
                updateD(station, 5, false);
                if (MasterSlaveTest) { updateD(slaveRow, 5, false); }

            }, cRunTest[station].Token); // end thread

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

    }
}