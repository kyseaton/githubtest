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

        private readonly object dataBaseLock = new object();

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
            
            // gettting rid of isASlave
            //you can't run tests from a slave anymore..
            //bool isASlave = false;
            // we will use this bool to say if we need to do slave stuff..
            bool MasterSlaveTest = false;
            int slaveRow = -1;

            //here for testing the not system
            sendNote(station,3,"Test Initiated");

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
                if ((string) d.Rows[station][1] == "")
                {
                    MessageBox.Show("Please Assign a Work Order");
                    return;
                }
                else if ((string) d.Rows[station][2] == "")
                {
                    MessageBox.Show("Please Select a Test Type.");
                    return;
                }
                else if ((bool)d.Rows[station][4] == false)
                {
                    MessageBox.Show("CScan is not In Use. Please Select it Before Proceeding.");
                    return;
                }
                else if (dataGridView1.Rows[station].Cells[4].Style.BackColor != Color.Green)
                {
                    MessageBox.Show("CScan is not currently connected.  Please Check Connection.");
                    return;
                }
                else if (GlobalVars.CScanData[station].cellCableType == "NONE")
                {
                    MessageBox.Show("CScan is not connected to a cells Cable.  Please connect a cells cable to this CSCAN to run a test.");
                    return;
                }

                // may also need to check some of this for the slave
                if (MasterSlaveTest)
                {
                    if ((string)d.Rows[slaveRow][1] == "")
                    {
                        MessageBox.Show("Please Assign a Work Order to the Slave Channel");
                        return;
                    }
                    else if ((bool)d.Rows[slaveRow][4] == false)
                    {
                        MessageBox.Show("The slave CScan is not In Use. Please make sure that is is In Use and connected before proceeding.");
                        return;
                    }
                    else if (dataGridView1.Rows[slaveRow].Cells[4].Style.BackColor != Color.Green)
                    {
                        MessageBox.Show("Slave CScan is not currently connected.  Please Check Connection.");
                        return;
                    }
                    else if (GlobalVars.CScanData[station].cellCableType == "NONE")
                    {
                        MessageBox.Show("Slave CScan is not connected to a cells Cable.  Please connect a cells cable to this CSCAN to run a test with it.");
                        return;
                    }
                }

                //I removed this because the charger present will determine how the test is run...
                // also need to check if an intelligent charger is connected for autoconfig
                //else if (GlobalVars.autoConfig == true && d.Rows[station][10].ToString().Contains("ICA") && (string)d.Rows[station][2] != "As Received")
                //{
                //    MessageBox.Show("Auto Configure is turned on, but there is no intelligent charger detected.  Please connect an intelligent charger or turn Auto Configure off in the tools menu.");
                //     return;
                //}

                else if (GlobalVars.autoConfig == true && (string)d.Rows[station][11] == "offline!" && (string)d.Rows[station][2] != "As Received" && d.Rows[station][10].ToString().Contains("ICA"))
                {
                    MessageBox.Show("Auto Configure is turned on, but the Intelligent Charger is set to be offline.  Please turn Auto Configure off in the tools menu or set the charger to be online by pressing the following key sequence on the charger: FUNC, 1, 1 and ENTER.");
                    return;
                }

                else if (((bool)d.Rows[station][8] == false || d.Rows[station][10].ToString() == "") && (string)d.Rows[station][2] != "As Received")
                {
                    // we don't have a charger linked. Do we still want to continue...
                    MessageBox.Show("There is no charger link established.  You mush have a charger or shunt to run any test other than 'As Received'");
                    return;
                }


                //Finally Lets check if the charger is running also...
                else if (d.Rows[station][11].ToString() == "RUN" && d.Rows[station][10].ToString().Contains("ICA"))
                {
                    //looks like that charger is already running.  Lets ask the user if we should stop the charger or not.
                    DialogResult dialogResult = MessageBox.Show("The charger appears to already be running. Do you want to stop it now and proceed with the test?", "Click Yes to have the program stop the charger or No to do it manually", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        criticalNum[Cstation] = true;
                        // now we need to reset the charger
                        updateD(station, 7, "Stopping Charger!");
                        // set KE1 to 2 ("command")
                        GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                        // set KE3 to stop
                        GlobalVars.ICSettings[Cstation].KE3 = (byte)2;
                        //Update the output string value
                        GlobalVars.ICSettings[Cstation].UpdateOutText();
                        //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                        for (int i = 0; i < 15; i++)
                        {
                            Thread.Sleep(1000);
                            if (GlobalVars.ICData[Cstation].runStatus != "RUN"){ break; }
                        }
                        // set KE1 to 1 ("query")
                        GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                        //Update the output string value
                        GlobalVars.ICSettings[Cstation].UpdateOutText();
                        criticalNum[Cstation] = false;
                    }
                    else
                    {
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
                    MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                    return;
                }
                
                #endregion



                // we passed the tests so we'll check the box to indicate we are a go!
                updateD(station, 5, true);
                if (MasterSlaveTest) { updateD(slaveRow, 5, true); }

                #region if we are doing the autoconfig, let's get the charger settings in order and then loaded into the charger! (case 2 only)

                // we'll tell the charger what to do! (if we have an IC and the user wants us to...)
                if (GlobalVars.autoConfig && d.Rows[station][10].ToString().Contains("ICA") && (string)d.Rows[station][2] != "As Received")
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
                        strAccessSelect = @"SELECT * FROM WorkOrders WHERE WorkOrderNumber='" + d.Rows[station][1].ToString() + "';";
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
                        byte[] tempKMStore = new byte[21] {48,48,48,48,48,48,48,48,48,48,48,48,48,48,48,48,48,48,48,48,48}; // 21 values for the 21 KM params

                        #region setting switch
                        try
                        {
                            switch ((string)d.Rows[station][2])
                            {
                                case "Full Charge-6":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][42].ToString().Substring(0, 2)));          //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][43].ToString()));                          //time hours
                                    //tempKMStore[2] = (byte)(48 + 1);                                                                            //time mins
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
                                    tempKMStore[8] = (byte)(48 + 1);                                                                            //time 2 mins
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

                                    tempKMStore[13] = (byte) 48;                                                                                //discharge time hours
                                    tempKMStore[14] = (byte) 48;                                                                                //discharge time mins
                                    tempKMStore[15] = (byte) 48;                                                                                //discharge current high byte
                                    tempKMStore[16] = (byte) 48;                                                                                //discharge current low byte
                                    tempKMStore[17] = (byte) 48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte) 48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte) 48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte) 48;                                                                                //discharge resistance low byte
                                    break;
                                case "Full Charge-4":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][52].ToString().Substring(0, 2)));          //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][53].ToString()));                          //time hours
                                    //tempKMStore[2] = (byte)(48 + 1);                                                                            //time mins
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
                                    tempKMStore[8] = (byte)(48 + 1);                                                                            //time 2 mins
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

                                    tempKMStore[13] = (byte) 48;                                                                                //discharge time hours
                                    tempKMStore[14] = (byte) 48;                                                                              //discharge time mins
                                    tempKMStore[15] = (byte) 48;                                                                                //discharge current high byte
                                    tempKMStore[16] = (byte) 48;                                                                                //discharge current low byte
                                    tempKMStore[17] = (byte) 48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte) 48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte) 48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte) 48;                                                                                //discharge resistance low byte
                                    break;
                                case "Top Charge-4":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][62].ToString().Substring(0, 2)));          //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][63].ToString()));                          //time hours
                                    tempKMStore[2] = (byte)(48 + 1);                                                                            //time mins
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

                                    tempKMStore[7] = (byte) 48;                                                                                 //time 2 hours
                                    tempKMStore[8] = (byte) 48;                                                                                 //time 2 mins
                                    tempKMStore[9] = (byte) 48;                                                                                 //top current 2 byte
                                    tempKMStore[10] = (byte) 48;                                                                                //bottom current 2 byte
                                    tempKMStore[11] = (byte) 48;                                                                                //top voltage 2 byte
                                    tempKMStore[12] = (byte) 48;                                                                                //bottom voltage 2 byte

                                    tempKMStore[13] = (byte) 48;                                                                                //discharge time hours
                                    tempKMStore[14] = (byte) 48;                                                                                //discharge time mins
                                    tempKMStore[15] = (byte) 48;                                                                                //discharge current high byte
                                    tempKMStore[16] = (byte) 48;                                                                                //discharge current low byte
                                    tempKMStore[17] = (byte) 48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte) 48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte) 48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte) 48;                                                                                //discharge resistance low byte
                                    break;
                                case "Top Charge-2":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][72].ToString().Substring(0, 2)));          //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][73].ToString()));                          //time hours
                                    tempKMStore[2] = (byte)(48 + 1);                                                                            //time mins
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

                                    tempKMStore[7] = (byte) 48;                                                                                 //time 2 hours
                                    tempKMStore[8] = (byte) 48;                                                                                 //time 2 mins
                                    tempKMStore[9] = (byte) 48;                                                                                 //top current 2 byte
                                    tempKMStore[10] = (byte) 48;                                                                                //bottom current 2 byte
                                    tempKMStore[11] = (byte) 48;                                                                                //top voltage 2 byte
                                    tempKMStore[12] = (byte) 48;                                                                                //bottom voltage 2 byte

                                    tempKMStore[13] = (byte) 48;                                                                                //discharge time hours
                                    tempKMStore[14] = (byte) 48;                                                                                //discharge time mins
                                    tempKMStore[15] = (byte) 48;                                                                                //discharge current high byte
                                    tempKMStore[16] = (byte) 48;                                                                                //discharge current low byte
                                    tempKMStore[17] = (byte) 48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte) 48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte) 48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte) 48;                                                                                //discharge resistance low byte
                                    break;
                                case "Top Charge-1":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][82].ToString().Substring(0, 2)));          //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][83].ToString()));                          //time hours
                                    tempKMStore[2] = (byte)(48 + 1);                                                                            //time mins
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

                                    tempKMStore[7] = (byte) 48;                                                                                 //time 2 hours
                                    tempKMStore[8] = (byte) 48;                                                                                 //time 2 mins
                                    tempKMStore[9] = (byte) 48;                                                                                 //top current 2 byte
                                    tempKMStore[10] = (byte) 48;                                                                                //bottom current 2 byte
                                    tempKMStore[11] = (byte) 48;                                                                                //top voltage 2 byte
                                    tempKMStore[12] = (byte) 48;                                                                                //bottom voltage 2 byte

                                    tempKMStore[13] = (byte) 48;                                                                                //discharge time hours
                                    tempKMStore[14] = (byte) 48;                                                                                //discharge time mins
                                    tempKMStore[15] = (byte) 48;                                                                                //discharge current high byte
                                    tempKMStore[16] = (byte) 48;                                                                                //discharge current low byte
                                    tempKMStore[17] = (byte) 48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte) 48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte) 48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte) 48;                                                                                //discharge resistance low byte
                                    break;
                                case "Capacity-1":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][92].ToString().Substring(0, 2)));          //mode
                                    tempKMStore[1] = (byte) 48;                                                                                 //time hours
                                    tempKMStore[2] = (byte) 48;                                                                                 //time mins
                                    tempKMStore[3] = (byte) 48;                                                                                 //top current byte
                                    tempKMStore[4] = (byte) 48;                                                                                 //bottom current byte
                                    tempKMStore[5] = (byte) 48;                                                                                 //top voltage byte
                                    tempKMStore[6] = (byte) 48;                                                                                 //bottom current byte

                                    tempKMStore[7] = (byte) 48;                                                                                 //time 2 hours
                                    tempKMStore[8] = (byte) 48;                                                                                 //time 2 mins
                                    tempKMStore[9] = (byte) 48;                                                                                 //top current 2 byte
                                    tempKMStore[10] = (byte) 48;                                                                                //bottom current 2 byte
                                    tempKMStore[11] = (byte) 48;                                                                                //top voltage 2 byte
                                    tempKMStore[12] = (byte) 48;                                                                                //bottom voltage 2 byte

                                    tempKMStore[13] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][93].ToString()));                         //discharge time hours
                                    tempKMStore[14] = (byte)(48 + 1);                                                                           //discharge time mins
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
                                    tempKMStore[1] = (byte) 48;                                                                                 //time hours
                                    tempKMStore[2] = (byte) 48;                                                                                 //time mins
                                    tempKMStore[3] = (byte) 48;                                                                                 //top current byte
                                    tempKMStore[4] = (byte) 48;                                                                                 //bottom current byte
                                    tempKMStore[5] = (byte) 48;                                                                                 //top voltage byte
                                    tempKMStore[6] = (byte) 48;                                                                                 //bottom current byte

                                    tempKMStore[7] = (byte) 48;                                                                                 //time 2 hours
                                    tempKMStore[8] = (byte) 48;                                                                                 //time 2 mins
                                    tempKMStore[9] = (byte) 48;                                                                                 //top current 2 byte
                                    tempKMStore[10] = (byte) 48;                                                                                //bottom current 2 byte
                                    tempKMStore[11] = (byte) 48;                                                                                //top voltage 2 byte
                                    tempKMStore[12] = (byte) 48;                                                                                //bottom voltage 2 byte

                                    tempKMStore[13] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][103].ToString()));                        //discharge time hours
                                    tempKMStore[14] = (byte)(48 + 1);                                                                           //discharge time mins
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
                                    tempKMStore[17] = (byte) 48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte) 48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte) 48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte) 48;                                                                                //discharge resistance low byte
                                    break;
                                case "Slow Charge-14":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][112].ToString().Substring(0, 2)));         //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][113].ToString()));                         //time hours
                                    tempKMStore[2] = (byte)(48 + 1);                         //time mins
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

                                    tempKMStore[7] = (byte) 48;                                                                                 //time 2 hours
                                    tempKMStore[8] = (byte) 48;                                                                                 //time 2 mins
                                    tempKMStore[9] = (byte) 48;                                                                                 //top current 2 byte
                                    tempKMStore[10] = (byte) 48;                                                                                //bottom current 2 byte
                                    tempKMStore[11] = (byte) 48;                                                                                //top voltage 2 byte
                                    tempKMStore[12] = (byte) 48;                                                                                //bottom voltage 2 byte

                                    tempKMStore[13] = (byte) 48;                                                                                //discharge time hours
                                    tempKMStore[14] = (byte) 48;                                                                                //discharge time mins
                                    tempKMStore[15] = (byte) 48;                                                                                //discharge current high byte
                                    tempKMStore[16] = (byte) 48;                                                                                //discharge current low byte
                                    tempKMStore[17] = (byte) 48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte) 48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte) 48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte) 48;                                                                                //discharge resistance low byte
                                    break;
                                case "Slow Charge-16":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][122].ToString().Substring(0, 2)));         //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][123].ToString()));                         //time hours
                                    tempKMStore[2] = (byte)(48 + 1);                         //time mins
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

                                    tempKMStore[7] = (byte) 48;                                                                                 //time 2 hours
                                    tempKMStore[8] = (byte) 48;                                                                                 //time 2 mins
                                    tempKMStore[9] = (byte) 48;                                                                                 //top current 2 byte
                                    tempKMStore[10] = (byte) 48;                                                                                //bottom current 2 byte
                                    tempKMStore[11] = (byte) 48;                                                                                //top voltage 2 byte
                                    tempKMStore[12] = (byte) 48;                                                                                //bottom voltage 2 byte

                                    tempKMStore[13] = (byte) 48;                                                                                //discharge time hours
                                    tempKMStore[14] = (byte) 48;                                                                                //discharge time mins
                                    tempKMStore[15] = (byte) 48;                                                                                //discharge current high byte
                                    tempKMStore[16] = (byte) 48;                                                                                //discharge current low byte
                                    tempKMStore[17] = (byte) 48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte) 48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte) 48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte) 48;                                                                                //discharge resistance low byte
                                    break;
                                case "Custom Chg":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][132].ToString().Substring(0, 2)));         //mode
                                    tempKMStore[1] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][133].ToString()));                         //time hours
                                    if (tempKMStore[0] != 20 && tempKMStore[0] != 21)
                                    {
                                        tempKMStore[2] = (byte)(48 + 1);                                                                        //time mins
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
                                        tempKMStore[8] = (byte)(48 + 1);                                                                        //time 2 mins
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

                                    tempKMStore[13] = (byte) 48;                                                                                //discharge time hours
                                    tempKMStore[14] = (byte) 48;                                                                                //discharge time mins
                                    tempKMStore[15] = (byte) 48;                                                                                //discharge current high byte
                                    tempKMStore[16] = (byte) 48;                                                                                //discharge current low byte
                                    tempKMStore[17] = (byte) 48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte) 48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte) 48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte) 48;                                                                                //discharge resistance low byte
                                    break;
                                case "Custom Cap":
                                    tempKMStore[0] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][142].ToString().Substring(0, 2)));         //mode
                                    tempKMStore[1] = (byte) 48;                                                                                 //time hours
                                    tempKMStore[2] = (byte) 48;                                                                                 //time mins
                                    tempKMStore[3] = (byte) 48;                                                                                 //top current byte
                                    tempKMStore[4] = (byte) 48;                                                                                 //bottom current byte
                                    tempKMStore[5] = (byte) 48;                                                                                 //top voltage byte
                                    tempKMStore[6] = (byte) 48;                                                                                 //bottom current byte

                                    tempKMStore[7] = (byte) 48;                                                                                 //time 2 hours
                                    tempKMStore[8] = (byte) 48;                                                                                 //time 2 mins
                                    tempKMStore[9] = (byte) 48;                                                                                 //top current 2 byte
                                    tempKMStore[10] = (byte) 48;                                                                                //bottom current 2 byte
                                    tempKMStore[11] = (byte) 48;                                                                                //top voltage 2 byte
                                    tempKMStore[12] = (byte) 48;                                                                                //bottom voltage 2 byte

                                    tempKMStore[13] = (byte)(48 + int.Parse(battery.Tables[0].Rows[0][143].ToString()));                        //discharge time hours
                                    tempKMStore[14] = (byte)(48 + 1);                                                                           //discharge time mins
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
                                    tempKMStore[2] = (byte)(48 + 1);                                                                            //time mins
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

                                    tempKMStore[7] = (byte) 48;                                                                                 //time 2 hours
                                    tempKMStore[8] = (byte) 48;                                                                                 //time 2 mins
                                    tempKMStore[9] = (byte) 48;                                                                                 //top current 2 byte
                                    tempKMStore[10] = (byte) 48;                                                                                //bottom current 2 byte
                                    tempKMStore[11] = (byte) 48;                                                                                //top voltage 2 byte
                                    tempKMStore[12] = (byte) 48;                                                                                //bottom voltage 2 byte

                                    tempKMStore[13] = (byte) 48;                                                                                //discharge time hours
                                    tempKMStore[14] = (byte) 48;                                                                                //discharge time mins
                                    tempKMStore[15] = (byte) 48;                                                                                //discharge current high byte
                                    tempKMStore[16] = (byte) 48;                                                                                //discharge current low byte
                                    tempKMStore[17] = (byte) 48;                                                                                //discharge voltage high byte
                                    tempKMStore[18] = (byte) 48;                                                                                //discharge voltage low byte
                                    tempKMStore[19] = (byte) 48;                                                                                //discharge resistance high byte
                                    tempKMStore[20] = (byte) 48;                                                                                //discharge resistance low byte
                                    break;
                                default:
                                    break;
                            }// end switch
                        }
                        catch
                        {
                            MessageBox.Show("Fail to pull the settings from the DataBase. \r\nPlease make sure you have the battery model setup for this test under the Manage Battery Models menu.");
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
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

                        //make sure the charger has priority
                        criticalNum[Cstation] = true;

                        Thread.Sleep(5000);
                        // set KE1 to 0 ("data")
                        GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                        GlobalVars.ICSettings[Cstation].UpdateOutText();

                        //turn off priority
                        criticalNum[Cstation] = false;

                    } // end try
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        MessageBox.Show("Error: Failed to Auto Load settings into the Charger.\n");
                        // reset everything
                        updateD(station, 5, false);
                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
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
                    strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME='"+ d.Rows[station][2].ToString() +"';";
                    DataSet settings = new DataSet();
                    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                    lock (dataBaseLock)
                    {
                        myAccessConn.Open();
                        myDataAdapter.Fill(settings, "TestType");
                        myAccessConn.Close();
                    }
                    
                    readings = (int) settings.Tables[0].Rows[0][3];
                    interval = (int) settings.Tables[0].Rows[0][4];

                }
                catch (Exception ex)
                {
                    myAccessConn.Close();
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    updateD(station, 5, false);
                    if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                    return;
                }
                #endregion

                int stepNum;
                int slaveStepNum = 0;

                //check if this is a new test first! Also check if the this is a master slave test and the two rows are out of sink...
                if ((string)d.Rows[station][6] == "" || (MasterSlaveTest && d.Rows[station][6].ToString() != d.Rows[slaveRow][6].ToString()))
                {

                    #region set up test number and ID
                    // Now we'll look up the current test number and increment the new step number/////////////////////////////////////////////////////
                    

                    //  open the db and pull in the options table
                    try
                    {
                        strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + d.Rows[station][1].ToString() + "' ORDER BY StepNumber DESC;";
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


                    }
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                        updateD(station, 5, false);
                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                        return;
                    }


                    //do we need to do the same for the slave?
                    if (MasterSlaveTest)
                    {
                        // Now we'll look up the current test number and increment the new step number/////////////////////////////////////////////////////
                        //  open the db and pull in the options table
                        try
                        {
                            strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + d.Rows[slaveRow][1].ToString() + "' ORDER BY StepNumber DESC;";
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


                        }
                        catch (Exception ex)
                        {
                            myAccessConn.Close();
                            MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                            return;
                        }
                    }




                    #endregion

                    // Save the test information to the test table///////////////////////////////////////////////////////////////////////////////////////////

                    // we need the technicial selected in the combo box too..
                    this.Invoke((MethodInvoker)delegate()
                    {
                        comboText = comboBox1.Text;
                    });

                    #region save new test to test table
                    //  now try to INSERT INTO it
                    try
                    {
                        string strUpdateCMD = "INSERT INTO Tests (WorkOrderId,WorkOrderNumber,AggrWorkOrders,StepNumber,[TestName],Readings,[Interval]," +
                            "StationNumber,Charger,DateStarted,DateCompleted,Technician,TerminalID,CellCableID,ShuntCableID,TempCableID,TerminalNumber," +
                            "Technology,CustomNoCells,BATNUMCABLE10) "
                            + "VALUES ('" +
                            "0" + "','" +                                                          //WorkOrderID (don't care..)
                            d.Rows[station][1].ToString().Trim() + "','" +                         //WorkOrderNumber
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
                            GlobalVars.CScanData[station].technology.ToString() + "','" +          //technology
                            GlobalVars.CScanData[station].customNoCells.ToString() + "','" +       //CustomNoCells
                            GlobalVars.CScanData[station].batNumCable10.ToString() +               //BATNUMCABLE10
                            "');";


                        OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                        lock (dataBaseLock)
                        {
                            myAccessConn.Open();
                            myAccessCommand.ExecuteNonQuery();
                            myAccessConn.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        MessageBox.Show("Error: Failed to store new data in the DataBase.\n" + ex.Message);
                        updateD(station, 5, false);
                        if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
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
                                "Technology,CustomNoCells,BATNUMCABLE10) "
                                + "VALUES ('" +
                                "0" + "','" +                                                          //WorkOrderID (don't care..)
                                d.Rows[slaveRow][1].ToString().Trim() + "','" +                         //WorkOrderNumber
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
                                GlobalVars.CScanData[slaveRow].technology.ToString() + "','" +          //technology
                                GlobalVars.CScanData[slaveRow].customNoCells.ToString() + "','" +       //CustomNoCells
                                GlobalVars.CScanData[slaveRow].batNumCable10.ToString() +               //BATNUMCABLE10
                                "');";


                            OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                            lock (dataBaseLock)
                            {
                                myAccessConn.Open();
                                myAccessCommand.ExecuteNonQuery();
                                myAccessConn.Close();
                            }

                        }
                        catch (Exception ex)
                        {
                            myAccessConn.Close();
                            MessageBox.Show("Error: Failed to store new data in the DataBase.\n" + ex.Message);
                            updateD(station, 5, false);
                            updateD(slaveRow, 5, false);
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
                    stepNum = int.Parse((string) d.Rows[station][3]);
                    if (MasterSlaveTest)
                    {
                        slaveStepNum = int.Parse((string)d.Rows[slaveRow][3]);
                    }
                }

                // and indicate that the test is starting
                updateD(station,7,"Starting Test");
                if(MasterSlaveTest){updateD(slaveRow,7,"Starting Test");}

                //reset the menu...
                this.Invoke((MethodInvoker)delegate()
                {
                    startNewTestToolStripMenuItem.Enabled = false;
                    resumeTestToolStripMenuItem.Enabled = false;
                    stopTestToolStripMenuItem.Enabled = true;
                });


                Thread.Sleep(1000);  // here so that we can actually see the grid update


                // OK now we'll tell the charger to startup (if we need to!)/////////////////////////////////////////////////////////////////////////////////////
                if ((string)d.Rows[station][2] == "As Received")
                {
                    // nothing to do! if it's an "As Received" or we are running the test on a slave charger... 
                }
                else if (d.Rows[station][10].ToString().Contains("ICA"))
                {
                    // we have an intelligent charger
                        //make sure the charger has priority
                        criticalNum[Cstation] = true;

                        // If we are in hold and we are starting a new test we need to reset before starting!
                        if ((string)d.Rows[station][11] != "RESET" && (string)d.Rows[station][6] == "")
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
                            //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                            for (int i = 0; i < 3; i++)
                            {
                                Thread.Sleep(1000);
                                if (GlobalVars.ICData[Cstation].runStatus != "HOLD") 
                                {
                                    break; 
                                }
                            }
                            updateD(station, 7, "Resetting Charger");
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Resetting Charger"); }
                            // set KE1 to 1 ("query")
                            GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                            //Update the output string value
                            GlobalVars.ICSettings[Cstation].UpdateOutText();
                            //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                            for (int i = 0; i < 15; i++)
                            {
                                Thread.Sleep(1000);
                                if (GlobalVars.ICData[Cstation].runStatus != "HOLD")
                                {
                                    break;
                                }
                            }

                        }

                        // we are not in hold and we had a fault that we needed to clear...
                        else if ((string)d.Rows[station][11] != "HOLD" && (string)d.Rows[station][6] != "")
                        {
                            // we are resuming after a fault has been corrected..
                            // now we need to reset the charger
                            updateD(station, 7, "Clearing Charger");
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Clearing Charger"); }
                            // set KE1 to 2 ("command")
                            GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                            // set KE3 to stop
                            GlobalVars.ICSettings[Cstation].KE3 = (byte)0;
                            //Update the output string value
                            GlobalVars.ICSettings[Cstation].UpdateOutText();
                            //now we are going to create a thread to set KE1 back to data mode after 5 seconds
                            Thread.Sleep(5000);
                            // set KE1 to 1 ("query")
                            GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                            //Update the output string value
                            GlobalVars.ICSettings[Cstation].UpdateOutText();

                        }
                    

                        updateD(station, 7, "Telling Charger to Run");
                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Telling Charger to Run"); }
                        // set KE1 to 2 ("command")
                        GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                        // set KE3 to run
                        GlobalVars.ICSettings[Cstation].KE3 = (byte)1;
                        //Update the output string value
                        GlobalVars.ICSettings[Cstation].UpdateOutText();
                        //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                        Thread.Sleep(5000);
                        // set KE1 to 1 ("query")
                        GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                        // set KE3 to 0 ("query")
                        GlobalVars.ICSettings[Cstation].KE3 = (byte)3;
                        //Update the output string value
                        GlobalVars.ICSettings[Cstation].UpdateOutText();
                        Thread.Sleep(5000);

                        //make sure the charger has priority
                        criticalNum[Cstation] = false;
                }// end else if for ICs...
                else if (d.Rows[station][10].ToString().Contains("CCA"))
                {
                    // We have a legacy Charger!
                    // We need to let it run!
                    GlobalVars.cHold[station] = false;
                }  // end else if for Legacy chargers
                else if (d.Rows[station][10].ToString().Contains("Shunt"))
                {
                    // We have a shunt!!!!!!
                    // We'll start the test when we start to see current...
                    updateD(station, 7, "Waiting!");
                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Waiting"); }
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
                                startNewTestToolStripMenuItem.Enabled = true;
                                resumeTestToolStripMenuItem.Enabled = true;
                                stopTestToolStripMenuItem.Enabled = false;
                            });

                            return;
                        }
                        #endregion
                        // look for a current
                        if (GlobalVars.CScanData[station].currentOne > 0.2)
                        {
                            // we found a current!
                            break;
                        }
                        Thread.Sleep(100);
                    }
                }  // end the shunt else!
                else
                {
                    // We don't have a charger linked ...
                    MessageBox.Show("Test failed!  Please check settings!");
                    return;
                }

                // We are now good to go on starting the test loop timer...
                // going to do the timming with a stop watch
                bool firstRun = true;  // so we know if we should call fillPlotCombos()
                int currentReading;
                string oldETime = "";
                var stopwatch = new Stopwatch();
                TimeSpan offset;

                //first check if we are resuming
                if ((string) d.Rows[station][6] != "")
                {
                    // we got a resume!
                    string temp = (string)d.Rows[station][6];
                    offset = new TimeSpan(int.Parse(temp.Substring(0, 2)), int.Parse(temp.Substring(3, 2)), int.Parse(temp.Substring(6, 2)));
                    currentReading = ((offset.Hours * 3600 + offset.Minutes * 60 + offset.Seconds) / interval) + 2;
                    updateD(station, 7, ("Reading " + (currentReading - 1).ToString() + " of " + readings.ToString()));
                    if (MasterSlaveTest) { updateD(slaveRow, 7, ("Reading " + (currentReading - 1).ToString() + " of " + readings.ToString())); }
                }
                else
                {
                    //fresh test!
                    offset = new TimeSpan();
                    currentReading = 1; 
                }

                TimeSpan eTime = new TimeSpan().Add(offset);
                string eTimeS = eTime.ToString(@"hh\:mm\:ss");
                stopwatch.Start();


                while (currentReading <= readings)
                {
                    // check if we need to take a reading
                    if (((currentReading - 1) * interval * 1000) < stopwatch.Elapsed.Add(offset).TotalMilliseconds )
                    {
                        //first record the elapsed amount of time
                        TimeSpan temp = stopwatch.Elapsed.Add(offset);
                        // update the grid
                        updateD(station,7,("Reading " + currentReading.ToString() + " of " + readings.ToString()));
                        if (MasterSlaveTest) { updateD(slaveRow, 7, ("Reading " + currentReading.ToString() + " of " + readings.ToString())); }

                        #region save a scan to the DB
                        //  now try to INSERT INTO it
                        try
                        {
                            string strUpdateCMD = "INSERT INTO ScanData (Station,BWO,STEP,RDG,[DATE],QS1,CTR,ETIME,CUR1,CUR2,VB1,VB2,VB3,VB4," +
                                "CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24," +
                                "BT1,BT2,BT3,BT4,BT5,BT6,CGND1,CGND2,REF,GND,FV,MSV,PSV) "
                                + "VALUES (" + station.ToString() + ",'" +                            //station number
                                d.Rows[station][1].ToString().Trim() + "','" +                          //WorkOrderNumber
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

                            //also insert the slave reading is need be...
                            if (MasterSlaveTest)
                            {

                                strUpdateCMD = "INSERT INTO ScanData (Station,BWO,STEP,RDG,[DATE],QS1,CTR,ETIME,CUR1,CUR2,VB1,VB2,VB3,VB4," +
                                    "CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24," +
                                    "BT1,BT2,BT3,BT4,BT5,BT6,CGND1,CGND2,REF,GND,FV,MSV,PSV) "
                                    + "VALUES (" + slaveRow.ToString() + ",'" +                            //slaveRow number
                                    d.Rows[slaveRow][1].ToString().Trim() + "','" +                          //WorkOrderNumber
                                    slaveStepNum.ToString("00") + "'," +                                            //StepNumber
                                    currentReading.ToString() + ",#" +                                     //ReadingNumber
                                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +                  //date
                                    GlobalVars.CScanData[slaveRow].QS1.ToString() + "','" +                  //QS1
                                    GlobalVars.CScanData[slaveRow].CTR.ToString() + "','" +                  //CTR
                                    temp.TotalDays.ToString("0.00000") + "','" +                                      //time elapsed in days
                                    GlobalVars.CScanData[slaveRow].currentOne.ToString("0.0") + "','" +           //CUR1
                                    GlobalVars.CScanData[slaveRow].currentTwo.ToString("0.0") + "','" +           //CUR2
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
                                    GlobalVars.CScanData[slaveRow].TP1.ToString("0.0") + "','" +                  //TP1
                                    GlobalVars.CScanData[slaveRow].TP2.ToString("0.0") + "','" +                  //TP2
                                    GlobalVars.CScanData[slaveRow].TP3.ToString("0.0") + "','" +                  //TP3
                                    GlobalVars.CScanData[slaveRow].TP4.ToString("0.0") + "','" +                  //TP4
                                    GlobalVars.CScanData[slaveRow].TP5.ToString("0.0") + "','" +                  //TP5
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
                                newRow["BWO"] = d.Rows[station][1].ToString().Trim();
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
                                newRow["BWO"] = d.Rows[slaveRow][1].ToString().Trim();
                                newRow["STEP"] = stepNum.ToString("00");
                                newRow["RDG"] = currentReading.ToString();
                                newRow["DATE"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                newRow["QS1"] = GlobalVars.CScanData[slaveRow].QS1.ToString();
                                newRow["CTR"] = GlobalVars.CScanData[slaveRow].CTR.ToString();
                                newRow["ETIME"] = temp.TotalDays.ToString("0.00000");
                                newRow["CUR1"] = GlobalVars.CScanData[slaveRow].currentOne.ToString("0.0");
                                newRow["CUR2"] = GlobalVars.CScanData[slaveRow].currentTwo.ToString("0.0");
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
                                newRow["BT1"] = GlobalVars.CScanData[slaveRow].TP1.ToString("0.0");
                                newRow["BT2"] = GlobalVars.CScanData[slaveRow].TP2.ToString("0.0");
                                newRow["BT3"] = GlobalVars.CScanData[slaveRow].TP3.ToString("0.0");
                                newRow["BT4"] = GlobalVars.CScanData[slaveRow].TP4.ToString("0.0");
                                newRow["BT5"] = GlobalVars.CScanData[slaveRow].TP5.ToString("0.0");
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
                            MessageBox.Show("Error: Failed to store new data in the DataBase.\n" + ex.Message);
                            updateD(station, 5, false);
                            if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
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
                            updateD(station,6,eTimeS);
                            if (MasterSlaveTest) { updateD(slaveRow, 6, eTimeS); }
                        }
                        catch { }
                        
                    }
                    oldETime = eTimeS;

                    #region Here is where wer are going to look for charging issues!
                    //Lets test for a charger issue now
                    // there are going to be three sections, IC section, CCA section and shunt section
                    if (d.Rows[station][10].ToString().Contains("ICA"))
                    {
                        if ((string)d.Rows[station][11] != "RUN" && (string)d.Rows[station][2] != "As Received")
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
                                    while ((string)d.Rows[station][11] != "HOLD")
                                    {
                                        //check for a cancel
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
                                                startNewTestToolStripMenuItem.Enabled = true;
                                                resumeTestToolStripMenuItem.Enabled = true;
                                                stopTestToolStripMenuItem.Enabled = false;
                                            });

                                            //return the charger to low priority
                                            criticalNum[Cstation] = false;

                                            return;
                                        }

                                        Thread.Sleep(400);
                                    }
                                    // were back!
                                    //start the charger back up!
                                    if ((string)d.Rows[station][9] != "" && (string)d.Rows[station][10] == "ICA" && (string)d.Rows[station][2] != "As Received")
                                    {
                                        //make sure the charger has priority
                                        criticalNum[Cstation] = true;

                                        updateD(station, 7, "Telling Charger to Run");
                                        if (MasterSlaveTest) { updateD(slaveRow, 7, "Telling Charger to Run"); }
                                        // set KE1 to 2 ("command")
                                        GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                                        // set KE3 to run
                                        GlobalVars.ICSettings[Cstation].KE3 = (byte)1;
                                        //Update the output string value
                                        GlobalVars.ICSettings[Cstation].UpdateOutText();
                                        //now we are going to create a thread to set KE1 back to data mode after 5 seconds
                                        Thread.Sleep(5000);
                                        // set KE1 to 1 ("query")
                                        GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                                        // set KE3 to 0 ("query")
                                        GlobalVars.ICSettings[Cstation].KE3 = (byte)3;
                                        //Update the output string value
                                        GlobalVars.ICSettings[Cstation].UpdateOutText();
                                        Thread.Sleep(10000);

                                        //make sure the charger no longer has priority
                                        criticalNum[Cstation] = false;
                                    }

                                    stopwatch.Start();
                                    updateD(station, 7, ("Reading " + currentReading.ToString() + " of " + readings.ToString()));
                                    if (MasterSlaveTest) { updateD(slaveRow, 7, ("Reading " + currentReading.ToString() + " of " + readings.ToString())); }

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
                                        startNewTestToolStripMenuItem.Enabled = true;
                                        resumeTestToolStripMenuItem.Enabled = true;
                                        stopTestToolStripMenuItem.Enabled = false;
                                    });

                                    //return the charger to low priority
                                    criticalNum[Cstation] = false;

                                    return;
                                }// end test fail else

                                //return the charger to low priority
                                criticalNum[Cstation] = false;
                            }  // end no run if
                        } // end IC block
                    }
                    else if (d.Rows[station][10].ToString().Contains("CCA"))
                    {

                    }
                    else if (d.Rows[station][10].ToString().Contains("Shunt"))
                    {

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
                        else if(d.Rows[station][10].ToString().Contains("ICA")){
                            //make sure the charger has priority
                            criticalNum[Cstation] = true;

                            // now we need to stop the charger
                            updateD(station, 7, "Telling Charger to Stop");
                            if (MasterSlaveTest) { updateD(slaveRow, 7, "Telling Charger to Stop"); }
                            // set KE1 to 2 ("command")
                            GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                            // set KE3 to stop
                            GlobalVars.ICSettings[Cstation].KE3 = (byte)2;
                            //Update the output string value
                            GlobalVars.ICSettings[Cstation].UpdateOutText();
                            //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                            Thread.Sleep(5000);
                            //turn off priority
                            criticalNum[Cstation] = false;
                        }
                        else if (d.Rows[station][10].ToString().Contains("CCA"))
                        {
                            // Put the charger back on hold...
                            GlobalVars.cHold[station] = true;
                        }
                        else if (d.Rows[station][10].ToString().Contains("Shunt"))
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
                            startNewTestToolStripMenuItem.Enabled = true;
                            resumeTestToolStripMenuItem.Enabled = true;
                            stopTestToolStripMenuItem.Enabled = false;
                        });

                        return;
                    }
                    #endregion

                    //every interval is defined in seconds to be safe, we'll test if we are at the correct interval every 200ms
                    Thread.Sleep(200);
                }

                // We finished so let's clearn up!
                // If we are running the charger tell it to stop and reset
                if ((string)d.Rows[station][9] != "" && (string)d.Rows[station][10] == "ICA" && (string)d.Rows[station][2] != "As Received")
                {
                    //make sure the charger has priority
                    criticalNum[Cstation] = true;

                    // now we need to stop the charger
                    updateD(station, 7, "Telling Charger to Stop");
                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Telling Charger to Stop"); }
                    // set KE1 to 2 ("command")
                    GlobalVars.ICSettings[Cstation].KE1 = (byte)2;
                    // set KE3 to stop
                    GlobalVars.ICSettings[Cstation].KE3 = (byte)2;
                    //Update the output string value
                    GlobalVars.ICSettings[Cstation].UpdateOutText();
                    //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                    Thread.Sleep(5000);
                    // now we need to reset the charger
                    updateD(station, 7, "Resetting Charger");
                    if (MasterSlaveTest) { updateD(slaveRow, 7, "Resetting Charger"); }
                    // set KE3 to RESET
                    GlobalVars.ICSettings[Cstation].KE3 = (byte)3;
                    //Update the output string value
                    GlobalVars.ICSettings[Cstation].UpdateOutText();
                    //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                    Thread.Sleep(5000);
                    // set KE1 to 1 ("query")
                    GlobalVars.ICSettings[Cstation].KE1 = (byte)0;
                    //Update the output string value
                    GlobalVars.ICSettings[Cstation].UpdateOutText();

                    //turn off priority
                    criticalNum[Cstation] = false;
                }
                

                //update the gui
                this.Invoke((MethodInvoker)delegate()
                {
                    startNewTestToolStripMenuItem.Enabled = true;
                    resumeTestToolStripMenuItem.Enabled = false;
                    stopTestToolStripMenuItem.Enabled = false;
                });

                //Test is finished!
                updateD(station,6,"");
                if (MasterSlaveTest) { updateD(slaveRow, 6, ""); }
                updateD(station,7,"Complete");
                if (MasterSlaveTest) { updateD(slaveRow, 7, "Complete"); }
                updateD(station, 5, false);
                if (MasterSlaveTest) { updateD(slaveRow, 5, false); }
                
            },cRunTest[station].Token); // end thread

        }// end RunTest

    }
}
