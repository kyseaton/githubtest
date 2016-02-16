using System;
using System.Diagnostics;
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

namespace NewBTASProto
{
    public partial class Main_Form : Form
        {
        
        //ComPort.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(port_DataReceived_1);
        //EventArgs E = new EventArgs();

        /// <summary>
        /// Serial Stuff defined here
        /// </summary>
        public SerialPort ICComPort = new SerialPort();

        public CancellationTokenSource cPollIC;

        //for checking the IC when selected
        bool check =  false;
        int toCheck;
        int chanNum;

        //for critical operations (Start,Stop, etc)
        public bool [] criticalNum = new bool[16];

        //Com Error count 
         byte[] comErrorNum = new byte[16] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
         byte[] comGoodNum = new byte[16] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

        public void pollICs()
        {

            //This code is here to poll the ICs/////////////////////////////////////////////////////////////////////////////////////
            // We are going to run all of this code on a helper thread, as to improve GUI performance//////////////////////////////////////////

            cPollIC = new CancellationTokenSource();

            string tempBuff = "";
            ICDataStore testData;

            // Open the comport
            ICComPort.ReadTimeout = 100;
            ICComPort.PortName = GlobalVars.ICComPort;
            ICComPort.BaudRate = 19200;
            ICComPort.DataBits = 8;
            ICComPort.StopBits = StopBits.One;
            ICComPort.Handshake = Handshake.None;
            ICComPort.Parity = Parity.None;
            ICComPort.DtrEnable = true;
            ICComPort.RtsEnable = false;
            ICComPort.Open();

            //OUTPUT DATA STRUCTURE ----------------------------------

            ThreadPool.QueueUserWorkItem(s =>
            {
                CancellationToken token = (CancellationToken)s;
                Thread.Sleep(1500);

                while (true)
                {
                    
                    try
                    {


                        for (int j = 0; j < 16; j++)
                        {
                            //sleep a little each time as to not overload the host
                            Thread.Sleep(10);
                            //putting the cancellation token in a often looked at place...
                            if (token.IsCancellationRequested) return;

                            int slaveRow = -1;
                            ////////////////////////////////////////////NORMAL PRIORITY CHARGERS ARE CHECKED HERE///////////////////////
                            if (//(bool)d.Rows[j][8] && 
                                (bool)d.Rows[j][4] && 
                                (string)d.Rows[j][9] != "" && 
                                !d.Rows[j][9].ToString().Contains("S"))
                            {
                                Thread.Sleep(800);
                                try
                                {
                                    // control for master slave setup
                                    int chargerID = 0;

                                    if (d.Rows[j][9].ToString().Length > 2)  // this is the case where we have a master and slave config
                                    {
                                        // we have a master slave charger
                                        // split into 3 and 4 digit case
                                        if (d.Rows[j][9].ToString().Length == 3)
                                        {
                                            // 3 case
                                            chargerID = int.Parse(d.Rows[j][9].ToString().Substring(0, 1));
                                            //now look for the slave row
                                            for (int i = 0; i < 16; i++)
                                            {
                                                if (d.Rows[i][9].ToString().Contains(chargerID.ToString()) && i != j ) //&& (bool) d.Rows[i][8])
                                                {
                                                    slaveRow = i;
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // 4 case
                                            chargerID = int.Parse(d.Rows[j][9].ToString().Substring(0, 2));
                                            for (int i = 0; i < 16; i++)
                                            {
                                                if (d.Rows[i][9].ToString().Contains(chargerID.ToString()) && i != j) //&& (bool)d.Rows[i][8])
                                                {
                                                    slaveRow = i;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    else  // this is the normal case with just one charger
                                    {
                                        chargerID = Convert.ToInt32(d.Rows[j][9]);
                                    }
                                    //Debug.Print("Normal Command Sent To " + chargerID.ToString());

                                    // send the short command based on the settings for the charger...
                                    ICComPort.Write(GlobalVars.ICSettings[chargerID].outText, 0, 28);
                                    // wait for a response
                                    tempBuff = ICComPort.ReadTo("Z");
                                    //do something with the new data
                                    char[] delims = { ' ' };
                                    string[] A = tempBuff.Split(delims);
                                    //A[1] has the terminal ID in it
                                    testData = new ICDataStore(A);
                                    GlobalVars.ICData[chargerID] = testData;

                                    //put this new data in the chart...
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        if (testData.online == true)
                                        {
                                            if ((bool)d.Rows[j][4]) //&& (bool)d.Rows[j][8]) // here to solve timing mismatch
                                            {
                                                if (testData.faultStatus != "") 
                                                {
                                                    updateD(j, 11, testData.faultStatus);
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow , 11, testData.faultStatus);
                                                    }
                                                }
                                                else if (testData.endStatus != "") 
                                                {
                                                    updateD(j, 11, testData.endStatus);
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 11, testData.endStatus);
                                                    }
                                                }
                                                else 
                                                { 
                                                    updateD(j, 11, testData.runStatus);
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 11, testData.runStatus);
                                                    }
                                                }
                                                if (dataGridView1.Rows[j].Cells[4].Style.BackColor != Color.Red)
                                                {
                                                    dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.YellowGreen;
                                                    if (slaveRow > -1)
                                                    {
                                                        dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.YellowGreen;
                                                    }
                                                }
                                            }
                                        }
                                        else if ((bool)d.Rows[j][4] ) //&& (bool)d.Rows[j][8]) // here to solve timing mismatch
                                        {
                                            updateD(j, 11, "offline!");
                                            dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Red;
                                            if (slaveRow > -1)
                                            {
                                                updateD(slaveRow, 11, "offline!");
                                                dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.Red;
                                            }
                                        }

                                        // also update the type of charger being used
                                        if ((bool)d.Rows[j][4] ) //&& (bool)d.Rows[j][8]) // here to solve timing mismatch
                                        {
                                            if (testData.boardID == 1) 
                                            { 
                                                updateD(j, 10, "ICA mini");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 10, "ICA mini");
                                                }
                                            }
                                            else if (testData.boardID == 6) 
                                            { 
                                                updateD(j, 10, "ICA SMC");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 10, "ICA SMC");
                                                }
                                            }
                                            else if (testData.boardID == 8) 
                                            { 
                                                updateD(j, 10, "ICA SMC ED");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 10, "ICA SMC ED");
                                                }
                                            }
                                            else if (testData.boardID == 4) 
                                            { 
                                                updateD(j, 10, "ICA SMI");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 10, "ICA SMI");
                                                }
                                            }
                                        }
                                        
                                        //rtbIncoming.Text = j.ToString() + "  :  " + tempBuff;
                                        
                                    });

                                    comErrorNum[j] = 0;

                                    Thread.Sleep(200);

                                }
                                catch (Exception ex)
                                {
                                    if (ex is System.TimeoutException)
                                    {
                                        if (comErrorNum[j] < 3) { comErrorNum[j]++; }
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            if (comErrorNum[j] > 2 && d.Rows[j][10].ToString().Contains("ICA"))
                                            {
                                                updateD(j, 11, "");
                                                updateD(j, 10, "");
                                                dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Gainsboro;
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 11, "");
                                                    updateD(slaveRow, 10, "");
                                                    dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.Gainsboro;
                                                }
                                            }
                                            tempBuff = ICComPort.ReadExisting();
                                            //rtbIncoming.Text = "Com Error" + System.Environment.NewLine + tempBuff;
                                        });
                                        Thread.Sleep(100);
                                    }
                                    else if (ex is System.IO.IOException || ex is System.InvalidOperationException)
                                    {
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            //let the user know that the comports are no longer working
                                            label8.Text = "Check Comports' Settings";
                                            label8.Visible = true;
                                            sendNote(0, 1, "COMPORTS DISCONNECTED. PLEASE CHECK.");
                                        });

                                        //cancel
                                        this.cPollIC.Cancel();
                                        this.cPollCScans.Cancel();
                                        this.sequentialScanT.Cancel();
                                        // close the com ports
                                        CSCANComPort.Close();
                                        ICComPort.Close();

                                        for (int q = 0; q < 16; q++)
                                        {
                                            cRunTest[q].Cancel();
                                        }
                                        return;
                                    }
                                    else if (ex is System.ObjectDisposedException)
                                    {
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            //let the user know that the comports are no longer working
                                            label8.Text = "Check Comports' Settings";
                                            label8.Visible = true;
                                            sendNote(0, 1, "COMPORTS DISCONNECTED. PLEASE CHECK.");
                                        });

                                        //cancel
                                        this.cPollIC.Cancel();
                                        this.cPollCScans.Cancel();
                                        this.sequentialScanT.Cancel();

                                        for (int q = 0; q < 16; q++)
                                        {
                                            cRunTest[q].Cancel();
                                        }

                                        return;
                                    }
                                    else { throw ex; }
                                }       // end catch
                            }       // end if

                            else if (d.Rows[j][10].ToString().Contains("ICA") ) // || (bool) d.Rows[j][8] == false) 
                            {
                                if ((string)d.Rows[j][11] != "" && !d.Rows[j][9].ToString().Contains("S"))
                                {
                                    updateD(j, 11, "");
                                    updateD(j, 10, "");
                                    dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Gainsboro;
                                    if (slaveRow > -1)
                                    {
                                        updateD(slaveRow, 11, "");
                                        updateD(slaveRow, 10, "");
                                        dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.Gainsboro;
                                    }

                                }
                              
                            }

                            /////////////////////////////CHECK FOR CHARGER IDENTITY///////////////////////////////////////
                            if (check)
                            {
                                Thread.Sleep(800);
                                try
                                {
                                    // send the short command based on the settings for the charger...
                                    ICComPort.Write(GlobalVars.ICSettings[toCheck].outText, 0, 28);
                                    //Debug.Print("Check Command Sent To " + toCheck.ToString());
                                    // wait for a response
                                    tempBuff = ICComPort.ReadTo("Z");
                                    //do something with the new data
                                    char[] delims = { ' ' };
                                    string[] A = tempBuff.Split(delims);
                                    //A[1] has the terminal ID in it
                                    testData = new ICDataStore(A);
                                    GlobalVars.ICData[toCheck] = testData;
                                    // if we got one then we can determine that we have an ICA
                                    if ((bool)d.Rows[chanNum][4] ) //&& (bool)d.Rows[chanNum][8]) // here to solve timing mismatch
                                    {
                                        if (testData.boardID == 1) { updateD(chanNum, 10, "ICA mini"); }
                                        else if (testData.boardID == 6) { updateD(chanNum, 10, "ICA SMC"); }
                                        else if (testData.boardID == 8) { updateD(chanNum, 10, "ICA SMC ED"); }
                                        else if (testData.boardID == 4) { updateD(chanNum, 10, "ICA SMI"); }
                                    }
                                    // and we don't need to check any more
                                    check = false;
                                    Thread.Sleep(200);
                                }
                                catch (Exception ex)
                                {
                                    if (ex is System.TimeoutException)
                                    {
                                        Thread.Sleep(100);
                                    }
                                    else if (ex is System.IO.IOException || ex is System.InvalidOperationException)
                                    {
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            //let the user know that the comports are no longer working
                                            label8.Text = "Check Comports' Settings";
                                            label8.Visible = true;
                                            sendNote(0, 1, "COMPORTS DISCONNECTED. PLEASE CHECK.");
                                        });

                                        //cancel
                                        this.cPollIC.Cancel();
                                        this.cPollCScans.Cancel();
                                        this.sequentialScanT.Cancel();

                                        for (int q = 0; q < 16; q++)
                                        {
                                            cRunTest[q].Cancel();
                                        }

                                        // close the com ports
                                        CSCANComPort.Close();
                                        ICComPort.Close();

                                        return;
                                    }
                                    else if (ex is System.ObjectDisposedException)
                                    {
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            //let the user know that the comports are no longer working
                                            label8.Text = "Check Comports' Settings";
                                            label8.Visible = true;
                                            sendNote(0, 1, "COMPORTS DISCONNECTED. PLEASE CHECK.");
                                        });

                                        //cancel
                                        this.cPollIC.Cancel();
                                        this.cPollCScans.Cancel();
                                        this.sequentialScanT.Cancel();

                                        for (int q = 0; q < 16; q++)
                                        {
                                            cRunTest[q].Cancel();
                                        }

                                        return;
                                    }
                                    else { throw ex; }
                                }       // end catch
                            }       // end else if

                            ////////////////////////////////////////////HIGH PRIORTIY CHARGERS ARE CHECKED HERE///////////////////////
                            // we need to check for critical operation also!
                            for(int i = 0;i < 16; i++)
                            {
                                slaveRow = -1;
                                if (criticalNum[i] == true)
                                {
                                    //Debug.Print("Station " + i.ToString() + " is critical");
                                    try
                                    {
                                        Thread.Sleep(800);
                                        // send the short command based on the settings for the charger...
                                        ICComPort.Write(GlobalVars.ICSettings[i].outText, 0, 28);
                                        //Debug.Print("High Priority Command Sent To " + i.ToString());
                                        // wait for a response
                                        tempBuff = ICComPort.ReadTo("Z");

                                        int station = 0;
                                        //find where the charger is located
                                        for (byte v = 0; v < 16; v++)
                                        {
                                            if (d.Rows[v][9].ToString() == "" || d.Rows[v][9].ToString().Contains("S") || !d.Rows[v][9].ToString().Contains(i.ToString())) { ;}  // do nothing if there is no assigned charger id or it's a slave id
                                            else if (d.Rows[v][9].ToString().Length > 2 )  // this is the case where we have a master and slave config
                                            {
                                                // we have a master slave charger
                                                // split into 3 and 4 digit case
                                                if (d.Rows[v][9].ToString().Length == 3)
                                                {
                                                    // 3 case
                                                    station = int.Parse(d.Rows[v][9].ToString().Substring(0,1));
                                                    //now look for the slave row
                                                    for (int ii = 0; ii < 16; ii++)
                                                    {
                                                        if (d.Rows[ii][9].ToString().Contains(v.ToString()) && ii != v ) //&& (bool) d.Rows[ii][8])
                                                        {
                                                            slaveRow = ii;
                                                            break;
                                                        }
                                                    }
                                                    break;  // found it!
                                                }
                                                else
                                                {
                                                    // 4 case
                                                    station = int.Parse(d.Rows[v][9].ToString().Substring(0, 2));
                                                    for (int ii = 0; ii < 16; ii++)
                                                    {
                                                        if (d.Rows[ii][9].ToString().Contains(v.ToString()) && ii != v ) //&& (bool)d.Rows[ii][8])
                                                        {
                                                            slaveRow = ii;
                                                            break;
                                                        }
                                                    }
                                                    break;  // found it!
                                                }
                                            }
                                            else if (int.Parse(d.Rows[v][9].ToString()) == i)  // this is the normal case of just one charger
                                            {
                                                station = v;
                                                break;
                                            }
                                        }


                                        // if we are running a test one repsonse may not be enough...
                                        if (comGoodNum[i] < 1 && (bool) d.Rows[i][5] == true)
                                        {
                                            comGoodNum[i]++;
                                            comErrorNum[i] = 0;
                                        }
                                        else
                                        {
                                            criticalNum[i] = false;
                                            comGoodNum[i] = 0;
                                            comErrorNum[i] = 0;
                                        }

                                        // we got a response so lets update the grid and the status box
                                        //A[1] has the terminal ID in it
                                        char[] delims = { ' ' };
                                        string[] A = tempBuff.Split(delims);
                                        testData = new ICDataStore(A);
                                        GlobalVars.ICData[station] = testData;
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            //rtbIncoming.Text = "Critical  " + i.ToString() + "  :  " + tempBuff;
                                            if (testData.online == true)
                                            {
                                                if ((bool)d.Rows[station][4] ) //&& (bool)d.Rows[station][8]) // here to solve timing mismatch
                                                {
                                                    if (testData.faultStatus != "") 
                                                    { 
                                                        updateD(station, 11, testData.faultStatus);
                                                        if (slaveRow > -1)
                                                        {
                                                            updateD(slaveRow, 11, testData.faultStatus);
                                                        }
                                                    }
                                                    else if (testData.endStatus != "") 
                                                    { 
                                                        updateD(station, 11, testData.endStatus);
                                                        if (slaveRow > -1)
                                                        {
                                                            updateD(slaveRow, 11, testData.endStatus);
                                                        }
                                                    }
                                                    else 
                                                    { 
                                                        updateD(station, 11, testData.runStatus);
                                                        if (slaveRow > -1)
                                                        {
                                                            updateD(slaveRow, 11, testData.runStatus);
                                                        }
                                                    }
                                                    if (dataGridView1.Rows[station].Cells[4].Style.BackColor != Color.Red)
                                                    {
                                                        dataGridView1.Rows[station].Cells[8].Style.BackColor = Color.YellowGreen;
                                                        if (slaveRow > -1)
                                                        {
                                                            dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.YellowGreen;
                                                        }
                                                    }
                                                }
                                            
                                            }
                                            else 
                                            {
                                                if ((bool)d.Rows[station][4] ) //&& (bool)d.Rows[station][8]) // here to solve timing mismatch
                                                {
                                                    updateD(station, 11, "offline!");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 11, "offline!");
                                                    }
                                                    //if ((bool)d.Rows[station][8]) 
                                                    //{ 
                                                        dataGridView1.Rows[station].Cells[8].Style.BackColor = Color.Red;
                                                        if (slaveRow > -1)
                                                        {
                                                            dataGridView1.Rows[station].Cells[8].Style.BackColor = Color.Red;
                                                        }
                                                    //}
                                                }
                            
                                            }


                                            if ((bool)d.Rows[station][4] ) //&& (bool)d.Rows[station][8]) // here to solve timing mismatch
                                            {
                                                // also update the type of charger being used
                                                if (testData.boardID == 1) 
                                                { 
                                                    updateD(station, 10, "ICA mini");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 10, "ICA mini");
                                                    }
                                                }
                                                else if (testData.boardID == 6) 
                                                { 
                                                    updateD(station, 10, "ICA SMC");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 10, "ICA SMC");
                                                    }
                                                }
                                                else if (testData.boardID == 8) 
                                                { 
                                                    updateD(station, 10, "ICA SMC ED");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 10, "ICA SMC ED");
                                                    }
                                                }
                                                else if (testData.boardID == 4) 
                                                { 
                                                    updateD(station, 10, "ICA SMI");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 10, "ICA SMI");
                                                    }
                                                }
                                            }

                                        });
                                        Thread.Sleep(200);
                                    }
                                    catch (Exception ex)
                                    {
                                        if (ex is System.TimeoutException)
                                        {
                                            //if this charger is not part of a test and we've had three errors, we need to turn off the critical...
                                            if (comErrorNum[i] < 3) 
                                            { 
                                                comGoodNum[i] = 0;
                                                comErrorNum[i]++; 
                                            }
                                            if ((bool) d.Rows[i][5] == false && comErrorNum[i] ==3)
                                            {
                                                criticalNum[i] = false;
                                                comErrorNum[i] = 0;
                                                comGoodNum[i] = 0;
                                            }
                                            Thread.Sleep(100);
                                        }
                                        else if (ex is System.IO.IOException || ex is System.InvalidOperationException)
                                        {
                                            this.Invoke((MethodInvoker)delegate
                                            {
                                                //let the user know that the comports are no longer working
                                                label8.Text = "Check Comports' Settings";
                                                label8.Visible = true;
                                                sendNote(0, 1, "COMPORTS DISCONNECTED. PLEASE CHECK.");
                                            });

                                            //cancel
                                            this.cPollIC.Cancel();
                                            this.cPollCScans.Cancel();
                                            this.sequentialScanT.Cancel();
                                            // close the com ports
                                            CSCANComPort.Close();
                                            ICComPort.Close();

                                            for (int q = 0; q < 16; q++)
                                            {
                                                cRunTest[q].Cancel();
                                            }

                                            return;
                                        }
                                        else if (ex is System.ObjectDisposedException)
                                        {
                                            this.Invoke((MethodInvoker)delegate
                                            {
                                                //let the user know that the comports are no longer working
                                                label8.Text = "Check Comports' Settings";
                                                label8.Visible = true;
                                                sendNote(0, 1, "COMPORTS DISCONNECTED. PLEASE CHECK.");
                                            });

                                            //cancel
                                            this.cPollIC.Cancel();
                                            this.cPollCScans.Cancel();
                                            this.sequentialScanT.Cancel();

                                            for (int q = 0; q < 16; q++)
                                            {
                                                cRunTest[q].Cancel();
                                            }

                                            return;
                                        }
                                        else { throw ex; }
                                    }       // end catch
                                } // end if
                            }// end for

                            ////////////////////////////////////////////MASTER FILLER DATA IS CHECKED HERE///////////////////////
                            if (GlobalVars.checkMasterFiller == true)
                            {
                                try
                                {
                                    Thread.Sleep(10);
                                    // send the short command to the masterfiller
                                    ICComPort.Write("~320Z");
                                    // wait for a response
                                    tempBuff = ICComPort.ReadTo("Z");

                                    // we got a response so lets update the grid and the status box
                                    //A[1] has the terminal ID in it
                                    char[] delims = { ' ' };
                                    string[] A = tempBuff.Split(delims);
                                    GlobalVars.MFData = A;
                                    GlobalVars.checkMasterFiller = false;
                                    Thread.Sleep(200);
                                }
                                catch (Exception ex)
                                {
                                    if (ex is System.TimeoutException)
                                    {
                                        // didn't get anything...
                                    }
                                    else if (ex is System.IO.IOException || ex is System.InvalidOperationException)
                                    {
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            //let the user know that the comports are no longer working
                                            label8.Text = "Check Comports' Settings";
                                            label8.Visible = true;
                                            sendNote(0, 1, "COMPORTS DISCONNECTED. PLEASE CHECK.");
                                        });

                                        //cancel
                                        this.cPollIC.Cancel();
                                        this.cPollCScans.Cancel();
                                        this.sequentialScanT.Cancel();
                                        // close the com ports
                                        CSCANComPort.Close();
                                        ICComPort.Close();

                                        for (int q = 0; q < 16; q++)
                                        {
                                            cRunTest[q].Cancel();
                                        }

                                        return;
                                    }
                                    else if (ex is System.ObjectDisposedException)
                                    {
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            //let the user know that the comports are no longer working
                                            label8.Text = "Check Comports' Settings";
                                            label8.Visible = true;
                                            sendNote(0, 1, "COMPORTS DISCONNECTED. PLEASE CHECK.");
                                        });

                                        //cancel
                                        this.cPollIC.Cancel();
                                        this.cPollCScans.Cancel();
                                        this.sequentialScanT.Cancel();

                                        for (int q = 0; q < 16; q++)
                                        {
                                            cRunTest[q].Cancel();
                                        }

                                        return;
                                    }
                                    else { throw ex; }
                                }       // end catch
                            } // end MasterFiller if

                        }           // end for
                    }               // end try
                    catch (Exception ex)
                    {
                        if (token.IsCancellationRequested) return;
                        else
                        {
                            //MessageBox.Show(ex.ToString());
                        }
                    }

                }                   // end while
            },cPollIC.Token); // end thread


        }  // end pollICs

        private void checkForIC(int chargerNum, int channelNum)
        {
            check = true;
            toCheck = chargerNum;
            chanNum = channelNum;
            // create a thread to check if we've got an IC connected at the selected address
             ThreadPool.QueueUserWorkItem(s =>
            {
                // we're going to give it 3 seconds to think about it...
                // it gets checked every other time...
                Thread.Sleep(12000);
                // now we'll make sure we're not looking anymore...
                check = false;
            }); // end thread


        }


    }
}
