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

        public CancellationTokenSource cRunTest;

        static System.Timers.Timer timer;

        //vars for recording in sc

        private void RunTest()
        {
            cRunTest = new CancellationTokenSource();
            int station = dataGridView1.CurrentRow.Index;

            // Everything is going to be done on a helper thread
            ThreadPool.QueueUserWorkItem(s =>
            {
                
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
                //else if (dataGridView1.Rows[station].Cells[4].Style.BackColor != Color.Green)
                //{
                //    MessageBox.Show("CScan is not currently connected.  Please Check Connection.");
                //    return;
                //}

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

                    myAccessConn.Open();
                    myDataAdapter.Fill(settings, "TestType");
                    readings = (int) settings.Tables[0].Rows[0][3];
                    interval = (int) settings.Tables[0].Rows[0][4];

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    myAccessConn.Close();
                    return;
                }

                // Now we'll look up the current test number and increment the new step number/////////////////////////////////////////////////////
                int stepNum;

                //  open the db and pull in the options table
                try
                {
                    strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + d.Rows[station][1].ToString() + "' ORDER BY StepNumber DESC;";
                    DataSet tests = new DataSet();
                    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                    myDataAdapter.Fill(tests, "Tests");
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
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    myAccessConn.Close();
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

                    myDataAdapter.Fill(testIDTable, "TestID");
                    testID = int.Parse(testIDTable.Tables[0].Rows[0][0].ToString()) + 1;

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    myAccessConn.Close();
                    return;
                }

                this.Invoke((MethodInvoker)delegate()
                {
                    comboText = comboBox1.Text;
                });

                // Save the test information to the test table///////////////////////////////////////////////////////////////////////////////////////////

                               
                                                                     
                //  now try to INSERT INTO it
                try
                {
                    string strUpdateCMD = "INSERT INTO Tests (TestID,WorkOrderId,WorkOrderNumber,AggrWorkOrders,StepNumber,[TestName],Readings,[Interval]," +
                        "StationNumber,Charger,DateStarted,DateCompleted,Technician,TerminalID,CellCableID,ShuntCableID,TempCableID,TerminalNumber," +
                        "Technology,CustomNoCells,BATNUMCABLE10) "
                        + "VALUES ('" + testID.ToString() + "','" +                  //TestID
                                                     "0" + "','" +                                 //WorkOrderID
                                                     d.Rows[station][1].ToString().Trim() + "','" +                   //WorkOrderNumber
                                                     "" + "','" +                                 //AggrWorkOrders
                                                     stepNum.ToString("00") + "','" +             //StepNumber
                                                     d.Rows[station][2].ToString() + "','" +      //TestName
                                                     readings.ToString() + "','" +                //Reading
                                                     (interval * 1000).ToString() + "','" +       //interval in msec
                                                     station.ToString() + "','" +                 // station number
                                                     d.Rows[station][10].ToString() + "',#" +      // charger type
                                                     DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,#" +            // start date
                                                     DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "#,'" +            // date completed
                        // do not enter! "##" + //"#,'" +                                 // restart date
                                                     comboText + "','" +                          // technician
                                                     (GlobalVars.CScanData[station].terminalID + 216).ToString() + "','" +    //terminal ID
                                                     GlobalVars.CScanData[station].CCID.ToString() + "','" +          //cells cable ID
                                                     GlobalVars.CScanData[station].SHCID.ToString() + "','" +         //shunt cable ID
                                                     GlobalVars.CScanData[station].TCAB.ToString() + "','" +          //temp cable ID
                                                     d.Rows[station][9].ToString() + "','" +      //charger ID (Terminal Number)
                                                     GlobalVars.CScanData[station].technology.ToString() + "','" +    //technology
                                                     GlobalVars.CScanData[station].customNoCells.ToString() + "','" + //CustomNoCells
                                                     GlobalVars.CScanData[station].batNumCable10.ToString() +// "','" + //BATNUMCABLE10
                                                     "');";  
                    OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to store new data in the DataBase.\n" + ex.Message);
                    myAccessConn.Close();
                    //return;
                }

                // We made it to this point without errors, so we'll update the grid with the step number
                lock(d)
                {
                    d.Rows[station][3] = stepNum.ToString();
                    d.Rows[station][7] = "Starting Test";
                }
                
                // also let the user know that we are starting the test
                

                // OK now we'll tell the charger to startup (if we need to!)/////////////////////////////////////////////////////////////////////////////////////

                if ((string)d.Rows[station][2] == "As Received")
                {
                    // we don't have to interact with the charger for this test
                }// end if
                else
                {
                    // we'll tell the charger what to do! (if we have an IC)
                    if (GlobalVars.autoConfig)
                    {
                        // change the settings of the charger before running it!

                    }


                    // we have an intelligent charger connected
                    if ((string)d.Rows[station][9] != "" && (string)d.Rows[station][10] == "ICA")
                    {
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
                TimeSpan eTime;
                string eTimeS = "";
                string oldETime = "";

                while (currentReading <= readings)
                {
                    // check if we need to take a reading
                    if (((currentReading - 1) * interval * 1000) < stopwatch.ElapsedMilliseconds )
                    {
                        d.Rows[station][7] = "Reading " + currentReading.ToString() + "of " + readings.ToString();
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
                            d.Rows[station][6] = eTimeS;
                        }
                        catch { }
                        
                    }
                    oldETime = eTimeS;
                    
                    //every interval is defined in seconds to be safe, we'll test if we are at the correct interval every 200ms
                    Thread.Sleep(200);
                }


                //Test is finished!
                d.Rows[station][6] = "";
                d.Rows[station][7] = "Complete";
                
            },cRunTest.Token); // end thread

        }// end RunTest

    }
}
