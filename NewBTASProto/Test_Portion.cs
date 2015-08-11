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
            int station = dataGridView1.CurrentRow.Index;
            cRunTest[station] = new CancellationTokenSource();

            // Everything is going to be done on a helper thread
            ThreadPool.QueueUserWorkItem(s =>
            {
                // setup the canellation token
                CancellationToken token = (CancellationToken)s;

                #region startup checks
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
                // also need to check if an intelligent charger is connected for autoconfig
                else if (GlobalVars.autoConfig == true && (string) d.Rows[station][10] != "ICA" && (string)d.Rows[station][2] != "As Received")
                {
                    MessageBox.Show("Auto Configure is turned on, but their is not Intelligent charger detected.  Please connect an intelligent charger or turn Auto Configure off in the tools menu.");
                    return;
                }
                else if (GlobalVars.autoConfig == true && (string) d.Rows[station][11] == "offline!" && (string)d.Rows[station][2] != "As Received")
                {
                    MessageBox.Show("Auto Configure is turned on, but the Intelligent Charger is set to be offline.  Please turn Auto Configure off in the tools menu or set the charger to be online by pressing the following key sequence on the charger: FUNC, 1, 1 and ENTER.");
                    return;
                }
                #endregion

                #region load test readings and interval values
                // Now we'll load the test parameters
                // We need to know the Interval and the number of readings///////////////////////////////////////////////////////////////////////

                int readings;
                int interval;

                string strAccessConn;
                string strAccessSelect;
                OleDbConnection myAccessConn;

                // create the connection
                try
                {
                    strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                    myAccessConn = new OleDbConnection(strAccessConn);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                    return;
                }

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
                    return;
                }
                #endregion

                #region set up test number and ID
                // Now we'll look up the current test number and increment the new step number/////////////////////////////////////////////////////
                int stepNum;

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
                    
                    if(tests.Tables[0].Rows.Count == 0)
                    {
                        stepNum = 1;                        
                    }
                    else
                    {
                        stepNum = int.Parse((string) tests.Tables[0].Rows[0][4]) + 1;
                    }
                    

                }
                catch (Exception ex)
                {
                    myAccessConn.Close();
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    return;
                }

                // we also need to come up with a new test ID///////////////////////////////////////////////////////////////////////////////////////
                int testID;

                try
                {
                    strAccessSelect = @"SELECT MAX(TestID) FROM Tests;";
                    DataSet testIDTable = new DataSet();
                    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                    lock (dataBaseLock)
                    {
                        myAccessConn.Open();
                        myDataAdapter.Fill(testIDTable, "TestID");
                        myAccessConn.Close();
                    }
                    
                    testID = int.Parse(testIDTable.Tables[0].Rows[0][0].ToString()) + 1;

                }
                catch (Exception ex)
                {
                    myAccessConn.Close();
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    return;
                }

                this.Invoke((MethodInvoker)delegate()
                {
                    comboText = comboBox1.Text;
                });

                #endregion

                // Save the test information to the test table///////////////////////////////////////////////////////////////////////////////////////////


                #region save new test to test table
                //  now try to INSERT INTO it
                try
                {
                    string strUpdateCMD = "INSERT INTO Tests (TestID,WorkOrderId,WorkOrderNumber,AggrWorkOrders,StepNumber,[TestName],Readings,[Interval]," +
                        "StationNumber,Charger,DateStarted,DateCompleted,Technician,TerminalID,CellCableID,ShuntCableID,TempCableID,TerminalNumber," +
                        "Technology,CustomNoCells,BATNUMCABLE10) "
                        + "VALUES ('" + testID.ToString() + "','" +                            //TestID
                        "0" + "','" +                                                          //WorkOrderID
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
                    return;
                }

                #endregion

                #region                // We made it to this point without errors, so we'll update the grid with the step number
                updateD(station,3,stepNum.ToString());
                // and indicate that the test is starting
                updateD(station,7,"Starting Test");
                // also make sure that the check box is checked
                updateD(station, 5, true);

                //reset the menu...
                this.Invoke((MethodInvoker)delegate()
                {
                    startNewTestToolStripMenuItem.Enabled = false;
                    stopTestToolStripMenuItem.Enabled = true;
                });


                Thread.Sleep(1000);  // here so that we can actually see the grid update

                #endregion

                // OK now we'll tell the charger to startup (if we need to!)/////////////////////////////////////////////////////////////////////////////////////

                if ((string)d.Rows[station][2] == "As Received" || GlobalVars.autoConfig == false)
                {
                    // we don't have to interact with the charger for this test
                }// end if
                else
                {
                    // open the DB
                    // we'll tell the charger what to do! (if we have an IC)
                    if (GlobalVars.autoConfig && (string) d.Rows[station][10] == "ICA")
                    {
                        // change the settings of the charger before running it!
                        
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
                            
                            string model = (string) workOrder.Tables[0].Rows[0][4];

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
                            
                            // now we can assigne the battery settings to the GlobalVars.
                            
                            // set KE1 to data
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KE1 = (byte)1;
                            // TODO pick a test mode based on the test setting

                            // update KM1
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KM1 = (byte)(10 + 48);

                            //// Charge Time 1
                            //Also Based on test choosen!!!!!!!!!!!!!!!
                            ////update KM2
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KM2 = (byte)(int.Parse((string) battery.Tables[0].Rows[0][13]) + 48);
                            ////update KM3
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KM3 = (byte)(0 + 48);

                            //// Charge Current 1
                            //// update KM4
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KM4 = (byte)(int.Parse((string)battery.Tables[0].Rows[0][12]) / 10 + 48);
                            ////update KM5
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KM5 = (byte)((int.Parse((string)battery.Tables[0].Rows[0][12]) % 10) * 10 + 48);

                            //// Charge Voltage 1
                            ////update KM6
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KM6 = (byte)(int.Parse((string)battery.Tables[0].Rows[0][14]) / 1 + 48);
                            ////update KM7
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KM7 = (byte)((int.Parse((string)battery.Tables[0].Rows[0][14]) % 1) * 100 + 48);
                            
                            //////////////////////////////////////////////////////////////////////////////////////////////////////

                            //// Charge Time 2
                            ////update KM8
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM8 = (byte)(numericUpDown8.Value + 48);
                            ////update KM9
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM9 = (byte)(numericUpDown7.Value + 48);

                            //// Charge Current 2
                            //// update KM10
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM10 = (byte)(numericUpDown6.Value / 10 + 48);
                            ////update KM11
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM11 = (byte)((numericUpDown6.Value % 10) * 10 + 48);

                            //// Charge Voltage 2
                            ////update KM12
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM12 = (byte)(numericUpDown5.Value / 1 + 48);
                            ////update KM13
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM13 = (byte)((numericUpDown5.Value % 1) * 100 + 48);

                            //////////////////////////////////////////////////////////////////////////////////////////////////////

                            //// Discharge Time
                            ////update KM14
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM14 = (byte)(numericUpDown12.Value + 48);
                            ////update KM15
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM15 = (byte)(numericUpDown11.Value + 48);

                            //// Discharge Current
                            //// update KM16
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM16 = (byte)(numericUpDown10.Value / 10 + 48);
                            ////update KM17
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM17 = (byte)((numericUpDown10.Value % 10) * 10 + 48);

                            //// Discharge Voltage
                            ////update KM18
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM18 = (byte)(numericUpDown9.Value / 1 + 48);
                            ////update KM19
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM19 = (byte)((numericUpDown9.Value % 1) * 100 + 48);

                            //// Discharge Resistance
                            ////update KM20
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM20 = (byte)(numericUpDown13.Value / 1 + 48);
                            ////update KM21
                            //GlobalVars.ICSettings[comboBox1.SelectedIndex].KM21 = (byte)((numericUpDown13.Value % 1) * 100 + 48);

                            //Update the output string value
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].UpdateOutText();
                            updateD(station, 7, "Loading Settings");

                            //make sure the charger has priority
                            criticalNum[station] = true;
            
                            Thread.Sleep(5000);
                            // set KE1 to 0 ("data")
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KE1 = (byte)0;
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].UpdateOutText();

                            //make sure the charger has priority
                            criticalNum[station] = false;

                        }
                        catch (Exception ex)
                        {
                            myAccessConn.Close();
                            MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                            return;
                        }

                    }


                    // we have an intelligent charger connected and it's not an as received test
                    if ((string)d.Rows[station][9] != "" && (string)d.Rows[station][10] == "ICA" && (string)d.Rows[station][2] != "As Received")
                    {
                        //make sure the charger has priority
                        criticalNum[station] = true;

                        updateD(station, 7, "Telling Charger to Run");
                        // set KE1 to 2 ("command")
                        GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KE1 = (byte)2;
                        // reset KE3
                        GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KE3 = (byte)1;
                        //Update the output string value
                        GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].UpdateOutText();
                        //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                        Thread.Sleep(5000);
                        // set KE1 to 1 ("query")
                        GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KE1 = (byte)0;
                        // set KE3 to 0 ("query")
                        GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KE3 = (byte)3;
                        //Update the output string value
                        GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].UpdateOutText();

                        //make sure the charger has priority
                        criticalNum[station] = false;
                    }
                    else
                    {
                        MessageBox.Show("Not Implemented for use without an intelligent charger!");
                    }
                }// end else

                // We are now good to go on starting the test loop timer...
                // going to do the timming with a stop watch
                var stopwatch = new Stopwatch();
                stopwatch.Start();
                int currentReading = 1;
                TimeSpan eTime = new TimeSpan();
                string eTimeS = eTime.ToString(@"hh\:mm\:ss");
                string oldETime = "";

                while (currentReading <= readings)
                {
                    // check if we need to take a reading
                    if (((currentReading - 1) * interval * 1000) < stopwatch.ElapsedMilliseconds )
                    {
                        //first record the elapsed amount of time
                        TimeSpan temp = stopwatch.Elapsed;
                        // update the grid
                        updateD(station,7,("Reading " + currentReading.ToString() + "of " + readings.ToString()));

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
                           

                        }
                        catch (Exception ex)
                        {
                            myAccessConn.Close();
                            MessageBox.Show("Error: Failed to store new data in the DataBase.\n" + ex.Message);
                            return;
                        }
                        #endregion

                        // finally update the reading
                        currentReading++;
                    }

                    //Now update the timer
                    eTime = TimeSpan.FromMilliseconds(stopwatch.ElapsedMilliseconds);
                    eTimeS = eTime.ToString(@"hh\:mm\:ss");
                    if (oldETime != eTimeS)
                    {
                        try
                        {
                            updateD(station,6,eTimeS);
                        }
                        catch { }
                        
                    }
                    oldETime = eTimeS;

                    //Now we should check for a cancel
                    #region cancel block
                    if (token.IsCancellationRequested)
                    {
                        if (GlobalVars.autoConfig == true)
                        {


                            //make sure the charger has priority
                            criticalNum[station] = true;

                            // now we need to stop the charger
                            updateD(station, 7, "Telling Charger to Stop");
                            // set KE1 to 2 ("command")
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KE1 = (byte)2;
                            // set KE3 to stop
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KE3 = (byte)2;
                            //Update the output string value
                            GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].UpdateOutText();
                            //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                            Thread.Sleep(5000);
                            //turn off priority
                            criticalNum[station] = false;
                        }
                        
                        //clear values from d
                        updateD(station, 7, ("Read " + currentReading.ToString() + "of " + readings.ToString()));
                        updateD(station, 5, false);

                        //update the gui
                        this.Invoke((MethodInvoker)delegate()
                        {
                            startNewTestToolStripMenuItem.Enabled = true;
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
                    criticalNum[station] = true;

                    // now we need to stop the charger
                    updateD(station, 7, "Telling Charger to Stop");
                    // set KE1 to 2 ("command")
                    GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KE1 = (byte)2;
                    // set KE3 to stop
                    GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KE3 = (byte)2;
                    //Update the output string value
                    GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].UpdateOutText();
                    //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                    Thread.Sleep(5000);
                    // now we need to reset the charger
                    updateD(station, 7, "Resetting Charger");
                    // set KE3 to RESET
                    GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KE3 = (byte)3;
                    //Update the output string value
                    GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].UpdateOutText();
                    //now we are going to create a thread to set KE1 back to data mode after 15 seconds
                    Thread.Sleep(5000);
                    // set KE1 to 1 ("query")
                    GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].KE1 = (byte)0;
                    //Update the output string value
                    GlobalVars.ICSettings[int.Parse((string)d.Rows[station][9])].UpdateOutText();

                    //turn off priority
                    criticalNum[station] = false;
                }
                

                //update the gui
                this.Invoke((MethodInvoker)delegate()
                {
                    startNewTestToolStripMenuItem.Enabled = true;
                    stopTestToolStripMenuItem.Enabled = false;
                });

                //Test is finished!
                updateD(station,6,"");
                updateD(station,7,"Complete");
                updateD(station, 5, false);
                
            },cRunTest[station].Token); // end thread

        }// end RunTest

    }
}
