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
        public SerialPort ICComPort;

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

         bool[] BBB_Toggle = new bool[16] { false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false};

        public void pollICs()
        {

            //This code is here to poll the ICs/////////////////////////////////////////////////////////////////////////////////////
            // We are going to run all of this code on a helper thread, as to improve GUI performance//////////////////////////////////////////

            cPollIC = new CancellationTokenSource();

            string tempBuff = "";
            ICDataStore testData;



            //OUTPUT DATA STRUCTURE ----------------------------------

            ThreadPool.QueueUserWorkItem(s =>
            {
                CancellationToken token = (CancellationToken)s;
                Thread.Sleep(1500);

                while (true)
                {
                    



                    for (int j = 0; j < 16; j++)
                    {
                        try
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
                                    if (ICComPort == null) 
                                    { 
                                    
                                    }
                                    // set up the comport
                                    ICComPort = new SerialPort();
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
                                    ICComPort.Write(GlobalVars.ICSettings[chargerID].outText, 0, 28);
                                    //Debug.Print(System.Text.Encoding.Default.GetString(GlobalVars.ICSettings[chargerID].outText));
                                    // wait for a response
                                    tempBuff = ICComPort.ReadTo("Z");
                                    ICComPort.Close();
                                    ICComPort.Dispose();
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
                                                else if (testData.availabilityStatus != "Enabled")
                                                {
                                                    //if the status is Bad Backup Batt we will do a send note, otherwise add it to the grid
                                                    if (testData.availabilityStatus == "Bad Backup Batt")
                                                    {
                                                        if (BBB_Toggle[j] == false)
                                                        {
                                                            BBB_Toggle[j] = true;
                                                            //do a send note!
                                                            this.Invoke((MethodInvoker)delegate()
                                                            {
                                                                sendNote(j, 3, "Bad Backup Batt");
                                                            });
                                                        }
                                                        updateD(j, 11, testData.runStatus);
                                                        if (slaveRow > -1)
                                                        {
                                                            updateD(slaveRow, 11, testData.runStatus);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        updateD(j, 11, testData.availabilityStatus);
                                                        if (slaveRow > -1)
                                                        {
                                                            updateD(slaveRow, 11, testData.availabilityStatus);
                                                        }
                                                    }
                                                }
                                                else 
                                                {
                                                    //clear the non critical fault status
                                                    BBB_Toggle[j] = false;
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
                                                updateD(j, 10, "ICA SMC EXD");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 10, "ICA SMC EXD");
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
                                            else if (testData.boardID == 20)
                                            {
                                                updateD(j, 10, "MFC-10");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 10, "MCF-10");
                                                }
                                            }
                                            else if (testData.boardID == 21)
                                            {
                                                updateD(j, 10, "MFC-25");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 10, "MCF-25");
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
                                            if (comErrorNum[j] > 2 && (d.Rows[j][10].ToString().Contains("ICA") || d.Rows[j][10].ToString().Contains("MFC")))
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
                                        ICComPort.Close();
                                        ICComPort.Dispose();
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
                                        CSCANComPort.Dispose();
                                        ICComPort.Close();
                                        ICComPort.Dispose();

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

                            else if ((d.Rows[j][10].ToString().Contains("ICA") || d.Rows[j][10].ToString().Contains("MFC"))) // || (bool) d.Rows[j][8] == false) 
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
                                    // set up the comport
                                    ICComPort = new SerialPort();
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
                                    ICComPort.Write(GlobalVars.ICSettings[toCheck].outText, 0, 28);
                                    //Debug.Print("Check Command Sent To " + toCheck.ToString());
                                    // wait for a response
                                    tempBuff = ICComPort.ReadTo("Z");
                                    ICComPort.Close();
                                    ICComPort.Dispose();
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
                                        else if (testData.boardID == 8) { updateD(chanNum, 10, "ICA SMC EXD"); }
                                        else if (testData.boardID == 4) { updateD(chanNum, 10, "ICA SMI"); }
                                        else if (testData.boardID == 20) { updateD(chanNum, 10, "MFC-10"); }
                                        else if (testData.boardID == 21) { updateD(chanNum, 10, "MFC-25"); }
                                    }
                                    // and we don't need to check any more
                                    check = false;
                                    Thread.Sleep(200);
                                }
                                catch (Exception ex)
                                {
                                    if (ex is System.TimeoutException)
                                    {
                                        ICComPort.Close();
                                        ICComPort.Dispose();
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
                                        CSCANComPort.Dispose();
                                        ICComPort.Close();
                                        ICComPort.Dispose();

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
                                        // set up the comport
                                        ICComPort = new SerialPort();
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
                                        ICComPort.Write(GlobalVars.ICSettings[i].outText, 0, 28);
                                        //Debug.Print("High Priority Command Sent To " + i.ToString());
                                        // wait for a response
                                        tempBuff = ICComPort.ReadTo("Z");
                                        ICComPort.Close();
                                        ICComPort.Dispose();

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
                                                    else if (testData.availabilityStatus != "Enabled")
                                                    {

                                                        //if the status is Bad Backup Batt we will do a send note, otherwise add it to the grid
                                                        if (testData.availabilityStatus == "Bad Backup Batt")
                                                        {
                                                            if (BBB_Toggle[station] == false)
                                                            {
                                                                BBB_Toggle[station] = true;
                                                                //do a send note!
                                                                this.Invoke((MethodInvoker)delegate()
                                                                {
                                                                    sendNote(station, 3, "Bad Backup Batt");
                                                                });
                                                            }
                                                            updateD(station, 11, testData.runStatus);
                                                            if (slaveRow > -1)
                                                            {
                                                                updateD(slaveRow, 11, testData.runStatus);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            updateD(station, 11, testData.availabilityStatus);
                                                            if (slaveRow > -1)
                                                            {
                                                                updateD(slaveRow, 11, testData.availabilityStatus);
                                                            }
                                                        }

                                                    }
                                                    else 
                                                    {
                                                        //clear the non critical fault status
                                                        BBB_Toggle[station] = false;
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
                                                    updateD(station, 10, "ICA SMC EXD");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 10, "ICA SMC EXD");
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
                                                else if (testData.boardID == 20)
                                                {
                                                    updateD(station, 10, "MFC-10");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 10, "MFC-10");
                                                    }
                                                }
                                                else if (testData.boardID == 20)
                                                {
                                                    updateD(station, 10, "MFC-25");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 10, "MFC-25");
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
                                            ICComPort.Close();
                                            ICComPort.Dispose();
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
                                            CSCANComPort.Dispose();
                                            ICComPort.Close();
                                            ICComPort.Dispose();

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
                                    // set up the comport
                                    ICComPort = new SerialPort();
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
                                    ICComPort.Write("~320Z");
                                    // wait for a response
                                    tempBuff = ICComPort.ReadTo("Z");
                                    ICComPort.Close();
                                    ICComPort.Dispose();
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
                                        ICComPort.Close();
                                        ICComPort.Dispose();
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
                                        CSCANComPort.Dispose();
                                        ICComPort.Close();
                                        ICComPort.Dispose();

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

                            ////////////////////////////////////////////AUTO SHORT BOARDS ARE SET HERE///////////////////////
                            #region AutoShort Comms

                            if (d.Rows[j][2].ToString().Contains("Shorting") && (bool)d.Rows[j][4] && GlobalVars.CScanData[j] != null && (GlobalVars.CScanData[j].CCID == 23 || GlobalVars.CScanData[j].CCID == 24)) //&& (bool) d.Rows[j][5] == true)
                            {
                                try
                                {
                                    // first lets create our output string

                                    string toAutoShort = "~" + (j + 32).ToString() + "L";
                                    //Create the 48 bit output
                                    char tempStore = (char)0;
                                    byte checkSum = 0;

                                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                                    for (int b = 0; b < 24; b++)
                                    {
                                        // only set bits if the test is running, otherwise clear everything by leaving them zero
                                        if ((bool)d.Rows[j][5] && b < GlobalVars.CScanData[j].cellsToDisplay)
                                        {
                                            if (Math.Abs(GlobalVars.CScanData[j].orderedCells[b]) < 0.1)
                                            {
                                                //set the 1 bit
                                                tempStore += (char)Math.Pow(2, ((2 * b) % 8));
                                            }
                                            if (Math.Abs(GlobalVars.CScanData[j].orderedCells[b]) < 1.5)
                                            {
                                                //set the 0 bit
                                                tempStore += (char)Math.Pow(2, ((2 * b + 1) % 8));
                                            }
                                        }

                                        // do we need to add the char to the output string?
                                        if((b+1) % 4 == 0)
                                        {
                                            sb.Append(tempStore);
                                            checkSum += (byte) tempStore;
                                            tempStore = (char) 0;
                                        }
                                    }


                                    toAutoShort += sb.ToString();
                                    toAutoShort += (char) checkSum;
                                    toAutoShort += "Z";

                                    Thread.Sleep(10);
                                    // send the short command to the autoshort box
                                    // set up the comport
                                    ICComPort = new SerialPort();
                                    ICComPort.Encoding = Encoding.GetEncoding(28591);
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
                                    ICComPort.Write(toAutoShort);
                                    // wait for a response
                                    tempBuff = ICComPort.ReadTo("Z");
                                    if (!tempBuff.Equals(toAutoShort.Substring(0, 11), StringComparison.Ordinal))
                                    {
                                        Debug.Print("toAutoShort");
                                        char[] chars = toAutoShort.ToCharArray();
                                        StringBuilder stringBuilder = new StringBuilder();
                                        foreach (char c in chars)
                                        {
                                            stringBuilder.Append(((Int16)c).ToString("x"));
                                        }
                                        String textAsHex = stringBuilder.ToString();
                                        Debug.Print(textAsHex);
                                    }
                                    ICComPort.Close();
                                    ICComPort.Dispose();
                                    // we got a response so lets update the grid and the status box
                                    //A[1] has the terminal ID in it

                                }
                                catch (Exception ex)
                                {
                                    if (ex is System.TimeoutException)
                                    {
                                        Debug.Print("time out");
                                        // didn't get anything...
                                        ICComPort.Close();
                                        ICComPort.Dispose();
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
                                        CSCANComPort.Dispose();
                                        ICComPort.Close();
                                        ICComPort.Dispose();

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
                                } // end catch
                            } // end main if
                        }               // end try
                        catch (Exception ex)
                        {
                            if (token.IsCancellationRequested) return;
                            else
                            {
                                //MessageBox.Show(ex.ToString());
                            }
                        }

#endregion

                    }           // end for


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
