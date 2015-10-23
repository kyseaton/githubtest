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
using System.Runtime.InteropServices;
using System.Data.OleDb;

namespace NewBTASProto
{
    public partial class Main_Form : Form
    {

        [DllImport("user32.dll")]
        static extern bool LockWindowUpdate(IntPtr hWndLock);

        //ComPort.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(port_DataReceived_1);
        //EventArgs E = new EventArgs();

        /// <summary>
        /// Serial Stuff defined here
        /// </summary>
        public SerialPort CSCANComPort = new SerialPort();

        internal delegate void SerialDataReceivedEventHandlerDelegate(object sender, SerialDataReceivedEventArgs e);
        delegate void SetTextCallback(string text);
        string InputData = String.Empty;

        // this is for the chart on the main form
        DataSet graphMainSet = new DataSet();

        // cancellation token
        public CancellationTokenSource cPollCScans;
        public CancellationTokenSource sequentialScanT;

        //Graph variables
        int technology1 = 0;
        int cell1 = 0;
        int type1 = 0;

        // prevent double fill combobox threads with the variable...
        // 99 is the start up value
        int oldRow = 99;

        //we are using this bool to say that we can go ahead and fill up the combo boxes
        //set it true when we get a good read
        //set if false otherwise

        bool goodRead = false;

        //to prevent plot gitter
        bool lockUpdate = false;


        public void pollCScans()
        {

            cPollCScans = new CancellationTokenSource();
            sequentialScanT = new CancellationTokenSource();

            string tempBuff;
            CScanDataStore testData;
            int tempClick = 0;

            //MOVE TO A STARTUP LOCATION!!!!!!!!!!!!!!!!!!!!!
            // Open the comport
            CSCANComPort.ReadTimeout = 200;
            CSCANComPort.PortName = GlobalVars.CSCANComPort;
            CSCANComPort.BaudRate = 19200;
            CSCANComPort.DataBits = 8;
            CSCANComPort.StopBits = StopBits.One;
            CSCANComPort.Handshake = Handshake.None;
            CSCANComPort.Parity = Parity.None;

            ThreadPool.QueueUserWorkItem(s =>
            {
                CancellationToken token = (CancellationToken)s;
                Thread.Sleep(1500);

                int slaveRow;

                while (true)
                {
                    
                    // this function consists of a while loop that is going to run until the thread is cancelled
                    try
                    {

                        for (int j = 0; j < 16; j++)
                        {


                            //pause for a little each time
                            Thread.Sleep(50);

                            // putting the cancel check in a well looked at place
                            if (token.IsCancellationRequested) return;

                            // first look at the selected row and then recheck all the other rows...
                            // look for the "In Use" columns and check for attached cscans
                            tempClick = dataGridView1.CurrentRow.Index;

                            // Selected Row Case////////////////////////////////////////////////////////////////////////
                            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][4])
                            {

                                try
                                {
                                    CSCANComPort.Open();

                                    // send the polling command
                                    string outText;
                                    int currentRow = dataGridView1.CurrentRow.Index;
                                    if (GlobalVars.cHold[currentRow])
                                    { outText = "~" + (currentRow + 16).ToString("00") + "L10Z"; }
                                    else
                                    { outText = "~" + (currentRow + 16).ToString("00") + "L00Z"; }
                                    CSCANComPort.Write(outText);
                                    // wait for a response
                                    
                                    tempBuff = CSCANComPort.ReadTo("Z");
                                    // close the comport
                                    CSCANComPort.Close();
                                    //we got a good read...
                                    goodRead = true;
                                    //do something with the new data
                                    char[] delims = { ' ' };
                                    string[] A = tempBuff.Split(delims);
                                    //A[1] has the terminal ID in it
                                    testData = new CScanDataStore(A);
                                    GlobalVars.CScanData[currentRow] = testData;

                                    if ((bool)d.Rows[currentRow][4] && tempClick == currentRow)  // test to see if we've clicked in the mean time...
                                    {

                                        //put this new data in the chart!
                                        this.Invoke((MethodInvoker)delegate
                                        {

                                            // first set the cell to green
                                            dataGridView1.Rows[currentRow].Cells[4].Style.BackColor = Color.Green;

                                            //update chart function
                                            updateChart(testData);



                                            //Real Time Data Portion
                                            string tempText = "";
                                            tempText = System.DateTime.Now.ToString("M/d/yyyy") + "       Terminal:  " + testData.terminalID.ToString() + Environment.NewLine;
                                            tempText += "Temp. Cable:  " + (3 - testData.TCAB).ToString() + "   (" + testData.tempPlateType + ")" + Environment.NewLine;
                                            tempText += "Cells Cable:  " + testData.CCID.ToString() + "   (" + testData.cellCableType + ")" + Environment.NewLine;
                                            tempText += "Shunt Cable:  " + testData.SHCID.ToString() + "   (" + testData.shuntCableType + ")" + Environment.NewLine;
                                            tempText += "Voltage Batt 1:  " + testData.VB1.ToString("00.00") + Environment.NewLine;
                                            tempText += "Voltage Batt 2:  " + testData.VB2.ToString("00.00") + Environment.NewLine;
                                            tempText += "Voltage Batt 3:  " + testData.VB3.ToString("00.00") + Environment.NewLine;
                                            tempText += "Voltage Batt 4:  " + testData.VB4.ToString("00.00") + Environment.NewLine;
                                            // select which currents to display
                                            if (testData.shuntCableType == "TEST BOX")
                                            {
                                                //dispaly both ...
                                                tempText += "Current#1:  " + testData.currentOne.ToString("00.00") + Environment.NewLine;
                                                tempText += "Current#2:  " + testData.currentTwo.ToString("00.00") + Environment.NewLine;
                                            }
                                            //if we have a mini that is charging...
                                            else if (d.Rows[currentRow][10].ToString().Contains("mini") && !(d.Rows[currentRow][2].ToString().Contains("Cap") || d.Rows[currentRow][2].ToString().Contains("Discharge")))
                                            {
                                                // for the mini case
                                                tempText += "Current:  " + testData.currentTwo.ToString("00.00") + Environment.NewLine;
                                            }
                                            else
                                            {
                                                // all other cases
                                                tempText += "Current:  " + testData.currentOne.ToString("00.00") + Environment.NewLine;
                                            }


                                            if (GlobalVars.Pos2Neg == false)
                                            {
                                                for (int i = 0; i < GlobalVars.CScanData[currentRow].cellsToDisplay; i++)
                                                {
                                                    tempText += "Cell #" + (i + 1).ToString() + ":  " + testData.orderedCells[i].ToString("0.000") + Environment.NewLine;
                                                }
                                            }
                                            else
                                            {
                                                for (int i = 0; i < GlobalVars.CScanData[currentRow].cellsToDisplay; i++)
                                                {
                                                    tempText += "Cell #" + (i + 1).ToString() + ":  " + testData.orderedCells[GlobalVars.CScanData[currentRow].cellsToDisplay - i - 1].ToString("0.000") + Environment.NewLine;
                                                }
                                            }

                                            // WE need to display open when we get -99, cold for -98, hot for -97 and shorted for -96
                                            switch (Convert.ToInt16(testData.TP1))
                                            {
                                                case -99:
                                                    tempText += "Temp Plate 1:  Open" + Environment.NewLine;
                                                    break;
                                                case -98:
                                                    tempText += "Temp Plate 1:  Cold" + Environment.NewLine;
                                                    break;
                                                case -97:
                                                    tempText += "Temp Plate 1:  Hot" + Environment.NewLine;
                                                    break;
                                                case -96:
                                                    tempText += "Temp Plate 1:  Shorted" + Environment.NewLine;
                                                    break;
                                                default:
                                                    tempText += "Temp Plate 1:  " + testData.TP1.ToString("00.0") + Environment.NewLine;
                                                    break;
                                            }
                                            switch (Convert.ToInt16(testData.TP2))
                                            {
                                                case -99:
                                                    tempText += "Temp Plate 2:  Open" + Environment.NewLine;
                                                    break;
                                                case -98:
                                                    tempText += "Temp Plate 2:  Cold" + Environment.NewLine;
                                                    break;
                                                case -97:
                                                    tempText += "Temp Plate 2:  Hot" + Environment.NewLine;
                                                    break;
                                                case -96:
                                                    tempText += "Temp Plate 2:  Shorted" + Environment.NewLine;
                                                    break;
                                                default:
                                                    tempText += "Temp Plate 2:  " + testData.TP2.ToString("00.0") + Environment.NewLine;
                                                    break;
                                            }
                                            switch (Convert.ToInt16(testData.TP3))
                                            {
                                                case -99:
                                                    tempText += "Temp Plate 3:  Open" + Environment.NewLine;
                                                    break;
                                                case -98:
                                                    tempText += "Temp Plate 3:  Cold" + Environment.NewLine;
                                                    break;
                                                case -97:
                                                    tempText += "Temp Plate 3:  Hot" + Environment.NewLine;
                                                    break;
                                                case -96:
                                                    tempText += "Temp Plate 3:  Shorted" + Environment.NewLine;
                                                    break;
                                                default:
                                                    tempText += "Temp Plate 3:  " + testData.TP3.ToString("00.0") + Environment.NewLine;
                                                    break;
                                            }
                                            switch (Convert.ToInt16(testData.TP4))
                                            {
                                                case -99:
                                                    tempText += "Temp Plate 4:  Open" + Environment.NewLine;
                                                    break;
                                                case -98:
                                                    tempText += "Temp Plate 4:  Cold" + Environment.NewLine;
                                                    break;
                                                case -97:
                                                    tempText += "Temp Plate 4:  Hot" + Environment.NewLine;
                                                    break;
                                                case -96:
                                                    tempText += "Temp Plate 4:  Shorted" + Environment.NewLine;
                                                    break;
                                                default:
                                                    tempText += "Temp Plate 4:  " + testData.TP1.ToString("00.0") + Environment.NewLine;
                                                    break;
                                            }
                                            switch (Convert.ToInt16(testData.TP5))
                                            {
                                                case -99:
                                                    tempText += "Ambient Temp:  Open" + Environment.NewLine;
                                                    break;
                                                case -98:
                                                    tempText += "Ambient Temp:  Cold" + Environment.NewLine;
                                                    break;
                                                case -97:
                                                    tempText += "Ambient Temp:  Hot" + Environment.NewLine;
                                                    break;
                                                case -96:
                                                    tempText += "Ambient Temp:  Shorted" + Environment.NewLine;
                                                    break;
                                                default:
                                                    tempText += "Ambient Temp:  " + testData.TP5.ToString("00.0") + Environment.NewLine;
                                                    break;
                                            }

                                            tempText += "Reference:  " + testData.ref95V.ToString("0.000") + Environment.NewLine;
                                            tempText += "Program Version " + testData.programVersion;

                                            LockWindowUpdate(label1.Handle);
                                            //label1.Text = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\n\rXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\r\nXXXXXXXXXXXXXXXXXX";
                                            //Thread.Sleep(2000);
                                            label1.Text = tempText;
                                            LockWindowUpdate(IntPtr.Zero);

                                        });

                                        ///////UPDATE CSCAN chargers here!////////////////////////////////////////////////////////

                                        currentRow = dataGridView1.CurrentRow.Index;

                                        // also look for a slave row
                                        slaveRow = -1;
                                        if (d.Rows[currentRow][9].ToString().Length > 2 && d.Rows[currentRow][9].ToString().Contains("M"))
                                        {
                                            // we have a master
                                            string temp = d.Rows[currentRow][9].ToString().Replace("-M", "");
                                            for (int q = 0; q < 16; q++)
                                            {
                                                if (d.Rows[q][9].ToString().Contains(temp) && d.Rows[q][9].ToString().Contains("S") && (bool) d.Rows[q][8])
                                                {
                                                    slaveRow = q;
                                                    break;
                                                }
                                            }
                                        }

                                        if ((bool)d.Rows[currentRow][4] &&
                                            (bool)d.Rows[currentRow][8] &&
                                            !d.Rows[currentRow][9].ToString().Contains("S") &&
                                            GlobalVars.CScanData[currentRow].connected &&
                                            (d.Rows[currentRow][10].ToString() == "" || dataGridView1.Rows[currentRow].Cells[8].Style.BackColor != Color.Olive || dataGridView1.Rows[currentRow].Cells[8].Style.BackColor != Color.Red))  // if a charger type isn't already there maybe we need to update with a CSCAN controlled charger...
                                        {
                                            // we got a CSCAN connected charger...
                                            updateD(currentRow, 10, "CCA");
                                            if (slaveRow > -1)
                                            {
                                                updateD(slaveRow, 10, "CCA");
                                            }
                                            if (GlobalVars.CScanData[currentRow].powerOn) 
                                            { 
                                                this.Invoke((MethodInvoker)delegate 
                                                {
                                                    dataGridView1.Rows[currentRow].Cells[8].Style.BackColor = Color.Olive;
                                                    if (slaveRow > -1)
                                                    {
                                                        dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.Olive;
                                                    }
                                                }); 
                                            }
                                            else
                                            {
                                                this.Invoke((MethodInvoker)delegate 
                                                { 
                                                    dataGridView1.Rows[currentRow].Cells[8].Style.BackColor = Color.Red;
                                                    if (slaveRow > -1)
                                                    {
                                                        dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.Red;
                                                    }
                                                });
                                                updateD(currentRow, 11, "Power Off");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 11, "Power Off");
                                                }
                                            }
                                        }
                                        //
                                        else if ((bool)d.Rows[currentRow][8] && GlobalVars.CScanData[currentRow].connected == false && !d.Rows[currentRow][9].ToString().Contains("S") && d.Rows[currentRow][10].ToString() == "CCA")
                                        {
                                            updateD(currentRow, 10, "");
                                            if (slaveRow > -1)
                                            {
                                                updateD(slaveRow, 10, "");
                                            }
                                            this.Invoke((MethodInvoker)delegate
                                            { 
                                                dataGridView1.Rows[currentRow].Cells[8].Style.BackColor = Color.Gainsboro;
                                                if (slaveRow > -1)
                                                {
                                                    dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.LightSteelBlue;
                                                }
                                            });
                                        }
                                        else if (dataGridView1.Rows[currentRow].Cells[8].Style.BackColor == Color.Olive && GlobalVars.CScanData[currentRow].powerOn == false && !d.Rows[currentRow][9].ToString().Contains("S") && d.Rows[currentRow][10].ToString() == "CCA")
                                        {
                                            this.Invoke((MethodInvoker)delegate 
                                            { 
                                                dataGridView1.Rows[currentRow].Cells[8].Style.BackColor = Color.Red;
                                                if (slaveRow > -1)
                                                {
                                                    dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.Red;
                                                }
                                            });
                                            updateD(currentRow, 11, "Power Off");
                                            if (slaveRow > -1)
                                            {
                                                updateD(slaveRow, 11, "Power Off");
                                            }
                                        }
                                        else if (dataGridView1.Rows[currentRow].Cells[8].Style.BackColor == Color.Red && GlobalVars.CScanData[currentRow].powerOn && !d.Rows[currentRow][9].ToString().Contains("S") && d.Rows[currentRow][10].ToString() == "CCA")
                                        {
                                            this.Invoke((MethodInvoker)delegate 
                                            { 
                                                dataGridView1.Rows[currentRow].Cells[8].Style.BackColor = Color.Olive;
                                                if (slaveRow > -1)
                                                {
                                                    dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.Olive;
                                                }
                                            });
                                        }



                                        ///////If nothing else do we at least have a shunt???!////////////////////////////////////////////////////////

                                        if ((bool)d.Rows[currentRow][4] && (bool)d.Rows[currentRow][8] && d.Rows[currentRow][10].ToString() == "" && !d.Rows[currentRow][9].ToString().Contains("S") && GlobalVars.CScanData[currentRow].shuntCableType != "NONE")
                                        {
                                            updateD(currentRow, 10, "Shunt");
                                            if (slaveRow > -1)
                                            {
                                                updateD(slaveRow, 10, "Shunt");
                                            }
                                            this.Invoke((MethodInvoker)delegate 
                                            { 
                                                dataGridView1.Rows[currentRow].Cells[8].Style.BackColor = Color.CadetBlue;
                                                if (slaveRow > -1)
                                                {
                                                    dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.CadetBlue;
                                                }
                                            });
                                        }
                                        else if ((bool)d.Rows[currentRow][4] && (bool)d.Rows[currentRow][8] && d.Rows[currentRow][10].ToString() == "Shunt" && !d.Rows[currentRow][9].ToString().Contains("S") && GlobalVars.CScanData[currentRow].shuntCableType != "NONE")
                                        {
                                            this.Invoke((MethodInvoker)delegate 
                                            { 
                                                dataGridView1.Rows[currentRow].Cells[8].Style.BackColor = Color.CadetBlue;
                                                if (slaveRow > -1)
                                                {
                                                    dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.CadetBlue;
                                                }
                                            });
                                        }
                                        else if ((bool)d.Rows[currentRow][4] && (bool)d.Rows[currentRow][8] && d.Rows[currentRow][10].ToString() == "Shunt" && !d.Rows[currentRow][9].ToString().Contains("S") && GlobalVars.CScanData[currentRow].shuntCableType == "NONE")
                                        {
                                            updateD(currentRow, 10, "");
                                            if (slaveRow > -1)
                                            {
                                                updateD(slaveRow, 10, "");
                                            }
                                            this.Invoke((MethodInvoker)delegate 
                                            { 
                                                dataGridView1.Rows[currentRow].Cells[8].Style.BackColor = Color.Gainsboro;
                                                if (slaveRow > -1)
                                                {
                                                    dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.LightSteelBlue;
                                                }
                                            });
                                        }

                                        /////////////Update the Run/Hold status of CCA chargers../////////////////////////////////////////////////////////
                                        if (d.Rows[currentRow][10].ToString().Contains("CCA") && !d.Rows[currentRow][9].ToString().Contains("S"))
                                        {
                                            if (GlobalVars.CScanData[currentRow].powerOn == false)
                                            {
                                                updateD(currentRow, 11, "Power Off");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 11, "Power Off");
                                                }
                                            }
                                            else
                                            {
                                                updateD(currentRow, 11, (GlobalVars.cHold[currentRow] ? "HOLD" : "RUN"));
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 11, (GlobalVars.cHold[currentRow] ? "HOLD" : "RUN"));
                                                }
                                            }
                                        }
                                        else if (d.Rows[currentRow][10].ToString().Contains("Shunt") && !d.Rows[currentRow][9].ToString().Contains("S"))
                                        {
                                            updateD(currentRow, 11, "");
                                            if (slaveRow > -1)
                                            {
                                                updateD(slaveRow, 11, "");
                                            }
                                        }

                                    }
                                }// end try
                                catch (Exception ex)
                                {
                                    if (ex is System.TimeoutException)
                                    {
                                        // make sure there haven't been any clicks in the mean time...
                                        if ((bool)d.Rows[tempClick][4] && tempClick == dataGridView1.CurrentRow.Index)
                                        {
                                            this.Invoke((MethodInvoker)delegate
                                            {
                                                if ((bool)d.Rows[tempClick][4] == true) { dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[4].Style.BackColor = Color.Red; }
                                                dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[8].Style.BackColor = Color.Gainsboro;
                                                chart1.Series.Clear();
                                                chart1.Invalidate();
                                                LockWindowUpdate(this.Handle);
                                                label1.Text = "";
                                                LockWindowUpdate(IntPtr.Zero);
                                                
                                            });
                                        }
                                        CSCANComPort.Close();

                                    }
                                }
                            }// end if for selected case
                            // the channel is not in use. clear everything!
                            else
                            {
                                this.Invoke((MethodInvoker)delegate
                                {
                                    chart1.Series.Clear();
                                    chart1.Invalidate();
                                    LockWindowUpdate(this.Handle);
                                    label1.Text = "";
                                    LockWindowUpdate(IntPtr.Zero);
                                });

                            }

                            /// NON Selected Row Case////////////////////////////////////////////////////////////////////////
                            // now look at all of the other cases to up date the label after a little break...
                            // if we are not looking for stations with the "find stations" function...

                            if (button1.Enabled == true)
                            {

                                if ((bool)d.Rows[j][4] && j != dataGridView1.CurrentRow.Index)
                                {

                                    // this allows for the current cscan being interrogated to be highlighted in the grid
                                    if (GlobalVars.highlightCurrent)
                                    {
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            dataGridView1.Rows[j].Cells[4].Style.BackColor = Color.Azure;
                                        });
                                    }


                                    // look at the "In Use" columns and check for attached cscans
                                    try
                                    {
                                        CSCANComPort.Open();
                                        // send the polling command
                                        string outText;
                                        if (GlobalVars.cHold[j])
                                        { outText = "~" + (j + 16).ToString("00") + "L10Z"; }
                                        else
                                        { outText = "~" + (j + 16).ToString("00") + "L00Z"; }
                                        CSCANComPort.Write(outText);

                                        // wait for a response
                                        tempBuff = CSCANComPort.ReadTo("Z");
                                        // close the comport
                                        CSCANComPort.Close();
                                        //do something with the new data
                                        char[] delims = { ' ' };
                                        string[] A = tempBuff.Split(delims);
                                        //A[1] has the terminal ID in it
                                        testData = new CScanDataStore(A);
                                        GlobalVars.CScanData[j] = testData;

                                        if ((bool)d.Rows[j][4])  // added to help with gui look
                                        {
                                            this.Invoke((MethodInvoker)delegate
                                            {
                                                // set the cell to green
                                                dataGridView1.Rows[j].Cells[4].Style.BackColor = Color.Green;
                                            });
                                        }  // end if

                                        ///////UPDATE CSCAN chargers here!////////////////////////////////////////////////////////

                                        // first look for a slave row
                                        // also look for a slave row
                                        slaveRow = -1;
                                        if (d.Rows[j][9].ToString().Length > 2 && d.Rows[j][9].ToString().Contains("M"))
                                        {
                                            string temp = d.Rows[j][9].ToString().Replace("-M", "");

                                            // we have a master
                                            for (int q = 0; q < 16; q++)
                                            {
                                                if (d.Rows[q][9].ToString().Contains(temp) && d.Rows[q][9].ToString().Contains("S") && (bool) d.Rows[q][8])
                                                {
                                                    slaveRow = q;
                                                    break;
                                                }
                                            }
                                        }

                                        if ((bool)d.Rows[j][4] && !d.Rows[j][9].ToString().Contains("S"))
                                        {
                                            if ((bool)d.Rows[j][4] &&
                                                (bool)d.Rows[j][8] &&
                                                GlobalVars.CScanData[j].connected &&
                                                (d.Rows[j][10].ToString() == "" || dataGridView1.Rows[j].Cells[8].Style.BackColor != Color.Olive || dataGridView1.Rows[j].Cells[8].Style.BackColor != Color.Red))  // if a charger type isn't already there maybe we need to update with a CSCAN controlled charger...
                                            {
                                                // we got a CSCAN connected charger...
                                                updateD(j, 10, "CCA");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 10, "CCA");
                                                }
                                                if (GlobalVars.CScanData[j].powerOn) 
                                                { 
                                                    this.Invoke((MethodInvoker)delegate 
                                                    {
                                                        dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Olive;
                                                        if (slaveRow > -1)
                                                        {
                                                            dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.Olive;
                                                        }
                                                    }); 
                                                }
                                                else 
                                                { 
                                                    this.Invoke((MethodInvoker)delegate 
                                                    { 
                                                        dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Red;
                                                        if (slaveRow > -1)
                                                        {
                                                            dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.Red;
                                                        }
                                                    });
                                                    updateD(j, 11, "Power Off");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 11, "Power Off");
                                                    }
                                                }
                                            }
                                            //
                                            else if ((bool)d.Rows[j][8] && GlobalVars.CScanData[j].connected == false && d.Rows[j][10].ToString() == "CCA" && !d.Rows[j][9].ToString().Contains("S"))
                                            {
                                                updateD(j, 10, "");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 10, "");
                                                }
                                                this.Invoke((MethodInvoker)delegate 
                                                { 
                                                    dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Gainsboro;
                                                    if (slaveRow > -1)
                                                    {
                                                        dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.LightSteelBlue;
                                                    }
                                                });
                                            }
                                            else if (dataGridView1.Rows[j].Cells[8].Style.BackColor == Color.Olive && GlobalVars.CScanData[j].powerOn == false && d.Rows[j][10].ToString() == "CCA" && !d.Rows[j][9].ToString().Contains("S"))
                                            {
                                                this.Invoke((MethodInvoker)delegate 
                                                { 
                                                    dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Red;
                                                    if (slaveRow > -1)
                                                    {
                                                        dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.Red;
                                                    }
                                                });

                                                updateD(j, 11, "Power Off");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 11, "Power Off");
                                                }
                                            }
                                            else if (dataGridView1.Rows[j].Cells[8].Style.BackColor == Color.Red && GlobalVars.CScanData[j].powerOn && d.Rows[j][10].ToString() == "CCA" && !d.Rows[j][9].ToString().Contains("S"))
                                            {
                                                this.Invoke((MethodInvoker)delegate 
                                                { 
                                                    dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Olive;
                                                    if (slaveRow > -1)
                                                    {
                                                        dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.Olive;
                                                    }
                                                });
                                            }
                                        }

                                        ///////If nothing else do we at least have a shunt???!////////////////////////////////////////////////////////

                                        if ((bool)d.Rows[j][4] && (bool)d.Rows[j][8] && d.Rows[j][10].ToString() == "" && GlobalVars.CScanData[j].shuntCableType != "NONE" && !d.Rows[j][9].ToString().Contains("S"))
                                        {
                                            updateD(j, 10, "Shunt");
                                            if (slaveRow > -1)
                                            {
                                                updateD(slaveRow, 10, "Shunt");
                                            }
                                            this.Invoke((MethodInvoker)delegate 
                                            { 
                                                dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.CadetBlue;
                                                if (slaveRow > -1)
                                                {
                                                    dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.CadetBlue;
                                                }
                                            });
                                        }
                                        else if ((bool)d.Rows[j][4] && (bool)d.Rows[j][8] && d.Rows[j][10].ToString() == "Shunt" && GlobalVars.CScanData[j].shuntCableType != "NONE" && !d.Rows[j][9].ToString().Contains("S"))
                                        {
                                            this.Invoke((MethodInvoker)delegate 
                                            { 
                                                dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.CadetBlue;
                                                if (slaveRow > -1)
                                                {
                                                    dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.CadetBlue;
                                                }
                                            });
                                        }
                                        else if ((bool)d.Rows[j][4] && (bool)d.Rows[j][8] && d.Rows[j][10].ToString() == "Shunt" && GlobalVars.CScanData[j].shuntCableType == "NONE" && !d.Rows[j][9].ToString().Contains("S"))
                                        {
                                            updateD(j, 10, "");
                                            if (slaveRow > -1)
                                            {
                                                updateD(slaveRow, 10, "");
                                            }
                                            this.Invoke((MethodInvoker)delegate 
                                            { 
                                                dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Gainsboro;
                                                if (slaveRow > -1)
                                                {
                                                    dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.LightSteelBlue;
                                                }
                                            });
                                        }

                                        /////////////Update the Run/Hold status of CCA chargers../////////////////////////////////////////////////////////
                                        if (d.Rows[j][10].ToString().Contains("CCA") && !d.Rows[j][9].ToString().Contains("S"))
                                        {
                                            if (GlobalVars.CScanData[j].powerOn == false)
                                            {
                                                updateD(j, 11, "Power Off");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 11, "Power Off");
                                                }
                                            }
                                            else
                                            {
                                                updateD(j, 11, (GlobalVars.cHold[j] ? "HOLD" : "RUN"));
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 11, (GlobalVars.cHold[j] ? "HOLD" : "RUN"));
                                                }
                                            }
                                            
                                        }
                                        else if (d.Rows[j][10].ToString().Contains("Shunt") && !d.Rows[j][9].ToString().Contains("S"))
                                        {
                                            updateD(j, 11, "");
                                            if (slaveRow > -1)
                                            {
                                                updateD(slaveRow, 11, "");
                                            }
                                        }
                                    }  // end try
                                    catch (Exception ex)
                                    {
                                        CSCANComPort.Close();
                                        if (ex is System.TimeoutException)
                                        {
                                            if ((bool)d.Rows[j][4] && dataGridView1.CurrentRow.Index != j)  // added to help with gui look
                                            {
                                                this.Invoke((MethodInvoker)delegate
                                                {
                                                    // set the cell to green
                                                    if ((bool)d.Rows[j][4]){dataGridView1.Rows[j].Cells[4].Style.BackColor = Color.Red;}
                                                    dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Gainsboro;
                                                });
                                            }
                                        }  // end if

                                    }  // end catch
                                }  // end if
                                else if ((bool)d.Rows[j][4] == false && dataGridView1.Rows[j].Cells[4].Style.BackColor != Color.Gainsboro && dataGridView1.Rows[j].Cells[4].Style.BackColor != Color.Empty)
                                {                                    
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                            dataGridView1.Rows[j].Cells[4].Style.BackColor = Color.Gainsboro; 
                                    });
                                }

                            }   // end if (this is the test to see if the find station function is running


                        }               // end for  this is the main for, which cycles throuhg the background channels
                    }                       // end try
                    catch (Exception ex)
                    {
                        if (token.IsCancellationRequested) return;
                        else
                        {
                            MessageBox.Show(ex.ToString());
                        }
                        
                    }                   // end catch
                }                       // end while (this is an endless loop, only the cancel token kills it)
            }, cPollCScans.Token);                     // end thread

        }

        private void updateChart(CScanDataStore testData)
        {
            //Replace based on values selected in radio1("Battery") or radio2 ("Cells")
            //and combo2 (Battery voltages) or combo3 (Cell voltages)
            //if that row is selected, update the chart portion
            int station = dataGridView1.CurrentRow.Index;

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            if (station != testData.terminalID || lockUpdate)
            {
                // we got bad input data due to multithreading...
                // or we are updating the combos..
                return;
            }
            //special 4 batt case
            else if ((comboBox2.Enabled == false || (radioButton2.Checked == true && comboBox3.Text == "Current Voltages")) && testData.CCID == 10)
            {
                //In this case we have a 4 Batt cable, but do not have a current test running.  We will display V1, V2, V3 and V4 on the main graph
                chart1.Series.Clear();

                var series1 = new System.Windows.Forms.DataVisualization.Charting.Series
                {
                    Name = "Series1",
                    Color = System.Drawing.Color.Green,
                    IsVisibleInLegend = false,
                    IsXValueIndexed = true,
                    ChartType = SeriesChartType.Column,
                    BorderColor = System.Drawing.Color.DarkGray,
                    BorderWidth = 1
                };
                this.chart1.Series.Add(series1);
                chart1.ChartAreas[0].AxisX.Title = "Battery";
                chart1.ChartAreas[0].AxisY.Title = "Voltage";

                int type = 4;

                if ((string)d.Rows[station][2] == "As Received" ||
                    (string)d.Rows[station][2] == "Capacity-1" ||
                    (string)d.Rows[station][2] == "Test" ||
                    (string)d.Rows[station][2] == "Custom Cap")
                {
                    type = 2;
                }
                else if ((string)d.Rows[station][2] == "Discharge")
                {
                    type = 0;
                }

                series1.Points.AddXY(1, testData.VB1);
                series1.Points[0].Color = pointColorMain(0, 1, testData.VB1, type);
                series1.Points[0].Label = "VB1";
                series1.Points.AddXY(2, testData.VB2);
                series1.Points[1].Color = pointColorMain(0, 1, testData.VB2, type);
                series1.Points[1].Label = "VB2";
                series1.Points.AddXY(3, testData.VB3);
                series1.Points[2].Color = pointColorMain(0, 1, testData.VB3, type);
                series1.Points[2].Label = "VB3";
                series1.Points.AddXY(4, testData.VB4);
                series1.Points[3].Color = pointColorMain(0, 1, testData.VB4, type);
                series1.Points[3].Label = "VB4";


                chart1.Invalidate();
                chart1.ChartAreas[0].RecalculateAxesScale();

            }
            //Normal Cell Voltage Only Case:
            else if (comboBox2.Enabled == false || (radioButton2.Checked == true && comboBox3.Text == "Cell Voltages"))
            {
                chart1.Series.Clear();
                
                var series1 = new System.Windows.Forms.DataVisualization.Charting.Series
                {
                    Name = "Series1",
                    Color = System.Drawing.Color.Green,
                    IsVisibleInLegend = false,
                    IsXValueIndexed = true,
                    ChartType = SeriesChartType.Column,
                    BorderColor = System.Drawing.Color.DarkGray,
                    BorderWidth = 1
                };
                this.chart1.Series.Add(series1);
                chart1.ChartAreas[0].AxisX.Title = "Cells";
                chart1.ChartAreas[0].AxisY.Title = "Voltage";

                int type = 4;

                if ((string)d.Rows[station][2] == "As Received" ||
                    (string)d.Rows[station][2] == "Capacity-1" ||
                    (string)d.Rows[station][2] == "Test" ||
                    (string)d.Rows[station][2] == "Custom Cap")
                {
                    type = 2;
                }
                else if ((string)d.Rows[station][2] == "Discharge")
                {
                    type = 0;
                }

                for (int i = 0; i < GlobalVars.CScanData[station].cellsToDisplay; i++)
                {
                    if (GlobalVars.Pos2Neg == false)
                    {
                        series1.Points.AddXY(i + 1, testData.orderedCells[i]);
                        // color test
                        series1.Points[i].Color = pointColorMain(0, 1, testData.orderedCells[i], type);
                    }
                    else
                    {
                        series1.Points.AddXY(i + 1, testData.orderedCells[GlobalVars.CScanData[station].cellsToDisplay - i - 1]);
                        // color test
                        series1.Points[i].Color = pointColorMain(0, 1, testData.orderedCells[GlobalVars.CScanData[station].cellsToDisplay - i - 1], type);
                    }
                }
                chart1.Invalidate();
                chart1.ChartAreas[0].RecalculateAxesScale();
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            else if (radioButton1.Checked == true)// we have to create a more complicated plot based on the values in GraphMain set...
            {
                //pulled this code from the Graphics_Form
                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //Battery Section!!!!
                try
                {
                    int q;
                    // only do something if the combo box has a test save!
                    if (comboBox2.SelectedIndex < 0) { return; }
                    // Here we will look at the Value selected and then plot graph1Set

                    //find out which graph to plot from the selected text
                    switch (comboBox2.Text)
                    {
                        case "Voltage":
                        case "Voltage 1":
                            q = 10;
                            chart1.ChartAreas[0].AxisY.Title = "Voltage";
                            break;
                        case "Voltage 2":
                            q = 11;
                            chart1.ChartAreas[0].AxisY.Title = "Voltage";
                            break;
                        case "Voltage 3":
                            q = 12;
                            chart1.ChartAreas[0].AxisY.Title = "Voltage";
                            break;
                        case "Voltage 4":
                            q = 13;
                            chart1.ChartAreas[0].AxisY.Title = "Voltage";
                            break;
                        case "Current":
                            q = 8;
                            chart1.ChartAreas[0].AxisY.Title = "Current";
                            break;
                        case "Temperature 1":
                            q = 38;
                            chart1.ChartAreas[0].AxisY.Title = "Temperature";
                            break;
                        case "Temperature 2":
                            q = 39;
                            chart1.ChartAreas[0].AxisY.Title = "Temperature";
                            break;
                        case "Temperature 3":
                            q = 40;
                            chart1.ChartAreas[0].AxisY.Title = "Temperature";
                            break;
                        case "Temperature 4":
                            q = 41;
                            chart1.ChartAreas[0].AxisY.Title = "Temperature";
                            break;
                        default:
                            q = 7;
                            chart1.ChartAreas[0].AxisY.Title = "Time";
                            break;
                    }

                    //we need to graph the col 7 as time and q as the value
                    this.chart1.Series.Clear();
                    var series1 = new System.Windows.Forms.DataVisualization.Charting.Series
                    {
                        Name = "Series1",
                        Color = System.Drawing.Color.Green,
                        IsVisibleInLegend = false,
                        IsXValueIndexed = true,
                        ChartType = SeriesChartType.Column,
                        BorderColor = System.Drawing.Color.DarkGray,
                        BorderWidth = 1
                    };
                    this.chart1.Series.Add(series1);
                    chart1.ChartAreas[0].AxisX.Title = "Time";

                    for (int i = 0; i < graphMainSet.Tables[0].Rows.Count; i++)
                    {
                        series1.Points.AddXY(Math.Round(double.Parse(graphMainSet.Tables[0].Rows[i][7].ToString()) * 1440), graphMainSet.Tables[0].Rows[i][q]);
                        // color test
                        series1.Points[i].Color = pointColorMain(technology1, cell1, double.Parse(graphMainSet.Tables[0].Rows[i][q].ToString()), type1);
                    }

                                            // pad with zero Vals to help with the look of the plot...
                        // first get the interval and total points
                        int interval = 1;
                        int points = 61;

                        switch (d.Rows[station][2].ToString())
                        {
                            case "As Received":
                                interval = 1 / 30;
                                points = 3;
                                break;
                            case "Full Charge-6":
                                interval = 5;
                                points = 73;
                                break;
                            case "Full Charge-4":
                                interval = 4;
                                points = 61;
                                break;
                            case "Top Charge-4":
                                interval = 4;
                                points = 61;
                                break;
                            case "Top Charge-2":
                                interval = 3;
                                points = 41;
                                break;
                            case "Top Charge-1":
                                interval = 1;
                                points = 61;
                                break;
                            case "Constant Voltage":
                                interval = 5;
                                points = 73;
                                break;
                            case "Capacity-1":
                                interval = 1;
                                points = 61;
                                break;
                            case "Discharge":
                                interval = 1;
                                points = 61;
                                break;
                            case "Slow Charge-14":
                                interval = 12;
                                points = 73;
                                break;
                            case "SlowCharge-16":
                                interval = 16;
                                points = 61;
                                break;
                            default:
                                //custom cap and charge get the default...
                                //Custom Chg
                                //Custom Cap
                                break;
                        }


                        if (graphMainSet.Tables[0].Rows.Count <= points-1)
                        {
                            for (int i = graphMainSet.Tables[0].Rows.Count; i <= points-1; i++)
                            {
                                series1.Points.AddXY(i*interval, 0);
                            }
                        }
            

                    chart1.Invalidate();
                    chart1.ChartAreas[0].RecalculateAxesScale();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                }
            }// end else  if

                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Cells Section!!!!
            else
            {
                try
                {
                    chart1.ChartAreas[0].AxisY.Title = "Voltage";
                    chart1.ChartAreas[0].AxisX.Title = "Time";
                    int q;
                    // only do something if the radio button is selected
                    if (radioButton2.Checked == false || comboBox3.SelectedIndex < 0) { return; }
                    // Here we will look at the Value selected and then plot graph1Set

                    //find out which graph to plot from the selected text
                    switch (comboBox3.Text)
                    {
                        case "Ending Voltages":
                            q = 999;
                            break;
                        case "Cell 1":
                            q = 14;
                            break;
                        case "Cell 2":
                            q = 15;
                            break;
                        case "Cell 3":
                            q = 16;
                            break;
                        case "Cell 4":
                            q = 17;
                            break;
                        case "Cell 5":
                            q = 18;
                            break;
                        case "Cell 6":
                            q = 19;
                            break;
                        case "Cell 7":
                            q = 20;
                            break;
                        case "Cell 8":
                            q = 21;
                            break;
                        case "Cell 9":
                            q = 22;
                            break;
                        case "Cell 10":
                            q = 23;
                            break;
                        case "Cell 11":
                            q = 24;
                            break;
                        case "Cell 12":
                            q = 25;
                            break;
                        case "Cell 13":
                            q = 26;
                            break;
                        case "Cell 14":
                            q = 27;
                            break;
                        case "Cell 15":
                            q = 28;
                            break;
                        case "Cell 16":
                            q = 29;
                            break;
                        case "Cell 17":
                            q = 30;
                            break;
                        case "Cell 18":
                            q = 31;
                            break;
                        case "Cell 19":
                            q = 32;
                            break;
                        case "Cell 20":
                            q = 33;
                            break;
                        case "Cell 21":
                            q = 34;
                            break;
                        case "Cell 22":
                            q = 35;
                            break;
                        case "Cell 23":
                            q = 36;
                            break;
                        case "Cell 24":
                            q = 37;
                            break;
                        default:
                            q = 999;
                            chart1.ChartAreas[0].AxisX.Title = "Cells";
                            break;
                    }

                    //we need to graph the col 7 as time and q as the value
                    this.chart1.Series.Clear();
                    var series1 = new System.Windows.Forms.DataVisualization.Charting.Series
                    {
                        Name = "Series1",
                        Color = System.Drawing.Color.Green,
                        IsVisibleInLegend = false,
                        IsXValueIndexed = true,
                        ChartType = SeriesChartType.Column,
                        BorderColor = System.Drawing.Color.DarkGray,
                        BorderWidth = 1
                    };
                    this.chart1.Series.Add(series1);

                    if (q == 999)
                    {
                        for (int i = 0; i < cell1; i++)
                        {
                            series1.Points.AddXY(i + 1, graphMainSet.Tables[0].Rows[graphMainSet.Tables[0].Rows.Count - 1][i + 14]);
                            // color test
                            series1.Points[i].Color = pointColorMain(technology1, 1, double.Parse(graphMainSet.Tables[0].Rows[graphMainSet.Tables[0].Rows.Count - 1][i + 14].ToString()), type1);
                        }
                    }
                    else
                    {
                        for (int i = 0; i < graphMainSet.Tables[0].Rows.Count; i++)
                        {
                            series1.Points.AddXY(Math.Round(double.Parse(graphMainSet.Tables[0].Rows[i][7].ToString()) * 1440), graphMainSet.Tables[0].Rows[i][q]);
                            // color test
                            series1.Points[i].Color = pointColorMain(technology1, 1, double.Parse(graphMainSet.Tables[0].Rows[i][q].ToString()), type1);
                        }
                        // pad with zero Vals to help with the look of the plot...
                        // first get the interval and total points
                        int interval = 1;
                        int points = 61;

                        switch (d.Rows[station][2].ToString())
                        {
                            case "As Received":
                                interval = 1 / 30;
                                points = 3;
                                break;
                            case "Full Charge-6":
                                interval = 5;
                                points = 73;
                                break;
                            case "Full Charge-4":
                                interval = 4;
                                points = 61;
                                break;
                            case "Top Charge-4":
                                interval = 4;
                                points = 61;
                                break;
                            case "Top Charge-2":
                                interval = 3;
                                points = 41;
                                break;
                            case "Top Charge-1":
                                interval = 1;
                                points = 61;
                                break;
                            case "Constant Voltage":
                                interval = 5;
                                points = 73;
                                break;
                            case "Capacity-1":
                                interval = 1;
                                points = 61;
                                break;
                            case "Discharge":
                                interval = 1;
                                points = 61;
                                break;
                            case "Slow Charge-14":
                                interval = 12;
                                points = 73;
                                break;
                            case "SlowCharge-16":
                                interval = 16;
                                points = 61;
                                break;
                            default:
                                //custom cap and charge get the default...
                                //Custom Chg
                                //Custom Cap
                                break;
                        }


                        if (graphMainSet.Tables[0].Rows.Count <= points-1)
                        {
                            for (int i = graphMainSet.Tables[0].Rows.Count; i <= points-1; i++)
                            {
                                series1.Points.AddXY(i*interval, 0);
                            }
                        }
                    }

                    chart1.Invalidate();
                    chart1.ChartAreas[0].RecalculateAxesScale();


                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                }
            }// end else


        }

        public void sequentialScan()
        {
            ThreadPool.QueueUserWorkItem(s =>
            {

                // set up thread
                int tempClick = 0;      // this var will store the last channels value between loops
                int multi = 0;
                CancellationToken token = (CancellationToken)s;
                Thread.Sleep(600);

                // here is the main loop
                while (true)
                {
                    // this function consists of a while loop that is going to run until the thread is cancelled
                    try         // on error the loop will just start again...
                    {
                        if (token.IsCancellationRequested) return;
                        Thread.Sleep(500);             // loop every 0.5 seconds
                        multi += 1;                    // increment multi
                        multi %= 10;                     // test every fourth count
                        if (checkBox1.Checked && multi == 0 && button1.Enabled == true)          // sequential scanning is turned on
                        {
                            tempClick = dataGridView1.CurrentRow.Index;
                            //search from tempclick onto the next "in use" row
                            for (int q = 1; q < 16; q++)
                            {
                                if ((bool)d.Rows[(tempClick + q) % 16][4])
                                {
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        dataGridView1.CurrentCell = dataGridView1.Rows[(tempClick + q) % 16].Cells[0];
                                        dataGridView1.ClearSelection();
                                    });
                                    break;
                                }
                            }  // end for
                        }// end if
                    }// end try
                    catch
                    {
                        // take a break and then start over...
                        Thread.Sleep(500);    
                    }
                }// end while
            }, sequentialScanT.Token);                     // end thread

        }

        private Color pointColorMain(int tech, int Cells, double Value, int type)
        {

            // normal vented NiCds
            double Min1 = 0.25;
            double Min2 = 1.5;
            double Min3 = 1.55;
            double Min4 = 1.7;
            double Max = 1.75;

            // special case for cable 10, sealed NiCds
            if (tech == 1)
            {
                Min1 = 0.25;
                Min2 = 1.45;
                Min3 = 1.5;
                Min4 = 1.65;
                Max = 1.7;
            }

            // these are the discharging cases...
            switch (type)
            {

                // these are the As Recieved color setting    
                case 1:
                    Min1 = 0.1;
                    Min2 = 1.2;
                    Min3 = 1.25;
                    Min4 = 1.25;
                    break;
                // these are the Discharge settings
                case 2:
                    Min1 = 0;
                    Min2 = 0.5;
                    Min3 = 0.5;
                    Min4 = 0.5;
                    break;
                // these are the Capacity
                case 3:
                    Min1 = 1;
                    Min2 = 1;
                    Min3 = 1.05;
                    Min4 = 11.7;
                    Max = 1.25;
                    break;
                // this is for charging Nicads
                case 4:
                    //Min1 = 0.1;
                    //Min2 = 1.2;
                    //Min3 = 1.25;
                    break;
                default:
                    break;
            }

            // scale the limits for the number of cells in the battery
            Min1 *= Cells;
            Min2 *= Cells;
            Min3 *= Cells;
            Min4 *= Cells;
            Max *= Cells;



            // with all of that said, let's start picking colors!
            if (type == 4)
            {
                if (Value > Min3) { return System.Drawing.Color.Red; }
                else if (Value > Min2) { return System.Drawing.Color.Orange; }
                else if (Value > Min1) { return System.Drawing.Color.Green; }
                else { return System.Drawing.Color.Blue; }
            }
            // this is for all charging operations not involving lead acid
            else if (tech != 2 && type == 0)
            {
                if (Value < Min2) { return System.Drawing.Color.Yellow; }
                else if (Value >= Min2 && Value < Min3) { return System.Drawing.Color.Orange; }
                else if (Value >= Min3 && Value < Min4) { return System.Drawing.Color.Green; }
                else if (Value >= Min4 && Value < Max) { return System.Drawing.Color.Blue; }
                else { return System.Drawing.Color.Red; }
            }
            // lead acid case
            else if (type == 0)
            {
                return System.Drawing.Color.Orange;
            }
            // this is for the Capacity test, Discharge and As Recieved
            else if (tech != 2)
            {
                if (Value < Min1) { return System.Drawing.Color.Red; }
                else if (Value < Min2) { return System.Drawing.Color.Yellow; }
                else if (Value < Min3) { return System.Drawing.Color.Orange; }
                else if (Value < Min4) { return System.Drawing.Color.Green; }
                else if (Value < Min2) { return System.Drawing.Color.Orange; }
            }
            //Finally the lead acid Capacity test, Discharge and As Recieved case
            else
            {
                if (Value > Cells * 1.75) { return System.Drawing.Color.Green; }
                else if (Value >= Cells * 1.67) { return System.Drawing.Color.Orange; }
                return System.Drawing.Color.Red;
            }

            return System.Drawing.Color.Green;


        }

        //stor the current test to compare to the old test
        string oldTest = "Nada";

        private void fillPlotCombos(int currentRow)
        {
            if (currentRow == oldRow && oldTest == d.Rows[currentRow][2].ToString())
            {
                return;
            }

            oldTest = d.Rows[currentRow][2].ToString();

            ThreadPool.QueueUserWorkItem(s =>
                {
                    lockUpdate = true;
                    // first things first
                    // if the cscan isn't in use then lets just set the drop downs to the default and return...
                    if ((bool) d.Rows[currentRow][4] == false)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            // just set to the cells readings..
                            comboBox2.Items.Clear();
                            comboBox3.Items.Clear();
                            radioButton1.Enabled = false;
                            radioButton2.Enabled = false;
                            updateR2(true);
                            comboBox2.Enabled = false;
                            comboBox3.Enabled = false;
                            radioButton2.Text = "Cells";
                            comboBox3.Items.Add("Cell Voltages");
                            updateC3("Cell Voltages");
                        });

                        oldRow = currentRow;
                        lockUpdate = false;
                        return;
                    }





                    //wait for a good read
                    for (int waitCount = 0; goodRead == false; waitCount++)
                    {
                        if (waitCount > 100) { break; }
                        Thread.Sleep(100);
                    }
                    string workOrder;
                    string testStep;

                    //this is here to stop double row fill operations..

                    oldRow = currentRow;

                    try
                    {
                        workOrder = d.Rows[currentRow][1].ToString();
                        testStep = d.Rows[currentRow][3].ToString();
                    }// end try
                    catch 
                    {
                        lockUpdate = false;
                        return; 
                    }
 
                    //make sure we have the info with which to act on...
                    if (workOrder == "" || testStep == "")
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            try
                            {
                                // just set to the cells readings..
                                comboBox2.Items.Clear();
                                comboBox3.Items.Clear();
                                radioButton1.Enabled = false;
                                radioButton2.Enabled = false;
                                updateR2(true);
                                comboBox2.Enabled = false;
                                comboBox3.Enabled = false;
                                if (GlobalVars.CScanData[currentRow] == null)
                                {
                                    radioButton2.Text = "Cells";
                                    comboBox3.Items.Add("Cell Voltages");
                                    updateC3("Cell Voltages");
                                }
                                else if (GlobalVars.CScanData[currentRow].CCID == 10)
                                {
                                    radioButton2.Text = "Cur Vs";
                                    comboBox3.Items.Add("Current Voltages");
                                    updateC3("Current Voltages");
                                }
                                else
                                {
                                    radioButton2.Text = "Cells";
                                    comboBox3.Items.Add("Cell Voltages");
                                    updateC3("Cell Voltages");
                                }
                            }// end try
                            catch { }// end catch
                        });
                        lockUpdate = false;
                    }
                    else
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            comboBox2.Enabled = false;
                            comboBox3.Enabled = false;
                        });

                        // do it on a helper thread!

                        // FIRST CLEAR THE OLD DATA SET!
                        graphMainSet.Clear();
                        // Open database containing all the battery data....
                        string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                        string strAccessSelect = @"SELECT * FROM ScanData WHERE BWO='" + workOrder + @"' AND STEP='" + Int32.Parse(testStep).ToString("00") + @"' ORDER BY RDG ASC";

                        //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
                        OleDbConnection myAccessConn = null;
                        // try to open the DB
                        try
                        {
                            myAccessConn = new OleDbConnection(strAccessConn);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                            lockUpdate = false;
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
                                myDataAdapter.Fill(graphMainSet, "ScanData");
                                myAccessConn.Close();
                            }


                        }
                        catch (Exception ex)
                        {
                            myAccessConn.Close();
                            MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                            lockUpdate = false;
                            return;
                        }


                        //we also need to figure out the type of battery being charged
                        // Open database containing all the battery data....
                        strAccessSelect = @"SELECT StepNumber,TestName, Technology, CustomNoCells FROM Tests WHERE WorkOrderNumber='" + workOrder + @"'";


                        DataSet testsPerformed = new DataSet();
                        myAccessConn = null;
                        // try to open the DB
                        try
                        {
                            myAccessConn = new OleDbConnection(strAccessConn);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                            lockUpdate = false;
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
                                myDataAdapter.Fill(testsPerformed, "Tests");
                                myAccessConn.Close();
                            }

                        }
                        catch (Exception ex)
                        {
                            myAccessConn.Close();
                            MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                            lockUpdate = false;
                            return;
                        }

                        //For the colors!!!!
                        try
                        {
                            technology1 = (int)testsPerformed.Tables["Tests"].Rows[1][2];
                            if (Int32.Parse(testsPerformed.Tables["Tests"].Rows[1][3].ToString()) != 0) { cell1 = Int32.Parse(testsPerformed.Tables["Tests"].Rows[1][3].ToString()); }
                            else { cell1 = GlobalVars.CScanData[currentRow].cellsToDisplay; }
                            // The final step is to update the type of test that was selected
                            if (d.Rows[currentRow][3].ToString().Contains("As Recieved")) { type1 = 1; }
                            else if (d.Rows[currentRow][3].ToString().Contains("Discharge")) { type1 = 2; }
                            else if (d.Rows[currentRow][3].ToString().Contains("Cap")) { type1 = 3; }
                            else { type1 = 0; }
                        }
                        catch
                        {
                            // didn't work
                            // just set the defaults...
                            this.Invoke((MethodInvoker)delegate()
                            {
                                // just set to the cells readings..
                                comboBox2.Items.Clear();
                                comboBox3.Items.Clear();
                                radioButton1.Enabled = false;
                                radioButton2.Enabled = false;
                                updateR2(true);
                                comboBox2.Enabled = false;
                                comboBox3.Enabled = false;
                                radioButton2.Text = "Cells";
                                comboBox3.Items.Add("Cell Voltages");
                                updateC3("Cell Voltages");
                            });

                            lockUpdate = false;
                            return;
                        }



                        string cellCable = GlobalVars.CScanData[currentRow].CCID.ToString();

                        this.Invoke((MethodInvoker)delegate()
                        {
                            switch (cellCable)
                            {
                                case "1":
                                    // Battery combobox
                                    comboBox2.Items.Clear();
                                    comboBox2.Items.Add("Voltage");
                                    comboBox2.Items.Add("Current");
                                    comboBox2.Items.Add("Temperature 1");
                                    comboBox2.Items.Add("Temperature 2");
                                    comboBox2.Items.Add("Temperature 3");
                                    comboBox2.Items.Add("Temperature 4");
                                    // Cells combobox
                                    comboBox3.Items.Clear();
                                    comboBox3.Items.Add("Cell Voltages");
                                    comboBox3.Items.Add("Cell 1");
                                    comboBox3.Items.Add("Cell 2");
                                    comboBox3.Items.Add("Cell 3");
                                    comboBox3.Items.Add("Cell 4");
                                    comboBox3.Items.Add("Cell 5");
                                    comboBox3.Items.Add("Cell 6");
                                    comboBox3.Items.Add("Cell 7");
                                    comboBox3.Items.Add("Cell 8");
                                    comboBox3.Items.Add("Cell 9");
                                    comboBox3.Items.Add("Cell 10");
                                    comboBox3.Items.Add("Cell 11");
                                    comboBox3.Items.Add("Cell 12");
                                    comboBox3.Items.Add("Cell 13");
                                    comboBox3.Items.Add("Cell 14");
                                    comboBox3.Items.Add("Cell 15");
                                    comboBox3.Items.Add("Cell 16");
                                    comboBox3.Items.Add("Cell 17");
                                    comboBox3.Items.Add("Cell 18");
                                    comboBox3.Items.Add("Cell 19");
                                    comboBox3.Items.Add("Cell 20");
                                    break;
                                case "3":
                                    // Battery combobox
                                    comboBox2.Items.Clear();
                                    comboBox2.Items.Add("Voltage 1");
                                    comboBox2.Items.Add("Voltage 2");
                                    comboBox2.Items.Add("Current");
                                    comboBox2.Items.Add("Temperature 1");
                                    comboBox2.Items.Add("Temperature 2");
                                    comboBox2.Items.Add("Temperature 3");
                                    comboBox2.Items.Add("Temperature 4");
                                    // Cells combobox
                                    comboBox3.Items.Clear();
                                    comboBox3.Items.Add("Cell Voltages");
                                    comboBox3.Items.Add("Cell 1");
                                    comboBox3.Items.Add("Cell 2");
                                    comboBox3.Items.Add("Cell 3");
                                    comboBox3.Items.Add("Cell 4");
                                    comboBox3.Items.Add("Cell 5");
                                    comboBox3.Items.Add("Cell 6");
                                    comboBox3.Items.Add("Cell 7");
                                    comboBox3.Items.Add("Cell 8");
                                    comboBox3.Items.Add("Cell 9");
                                    comboBox3.Items.Add("Cell 10");
                                    comboBox3.Items.Add("Cell 11");
                                    comboBox3.Items.Add("Cell 12");
                                    comboBox3.Items.Add("Cell 13");
                                    comboBox3.Items.Add("Cell 14");
                                    comboBox3.Items.Add("Cell 15");
                                    comboBox3.Items.Add("Cell 16");
                                    comboBox3.Items.Add("Cell 17");
                                    comboBox3.Items.Add("Cell 18");
                                    comboBox3.Items.Add("Cell 19");
                                    comboBox3.Items.Add("Cell 20");
                                    comboBox3.Items.Add("Cell 21");
                                    comboBox3.Items.Add("Cell 22");
                                    break;
                                case "4":
                                    // Battery combobox
                                    comboBox2.Items.Clear();
                                    comboBox2.Items.Add("Voltage 1");
                                    comboBox2.Items.Add("Voltage 2");
                                    comboBox2.Items.Add("Voltage 3");
                                    comboBox2.Items.Add("Current");
                                    comboBox2.Items.Add("Temperature 1");
                                    comboBox2.Items.Add("Temperature 2");
                                    comboBox2.Items.Add("Temperature 3");
                                    comboBox2.Items.Add("Temperature 4");
                                    // Cells combobox
                                    comboBox3.Items.Clear();
                                    comboBox3.Items.Add("Cell Voltages");
                                    comboBox3.Items.Add("Cell 1");
                                    comboBox3.Items.Add("Cell 2");
                                    comboBox3.Items.Add("Cell 3");
                                    comboBox3.Items.Add("Cell 4");
                                    comboBox3.Items.Add("Cell 5");
                                    comboBox3.Items.Add("Cell 6");
                                    comboBox3.Items.Add("Cell 7");
                                    comboBox3.Items.Add("Cell 8");
                                    comboBox3.Items.Add("Cell 9");
                                    comboBox3.Items.Add("Cell 10");
                                    comboBox3.Items.Add("Cell 11");
                                    comboBox3.Items.Add("Cell 12");
                                    comboBox3.Items.Add("Cell 13");
                                    comboBox3.Items.Add("Cell 14");
                                    comboBox3.Items.Add("Cell 15");
                                    comboBox3.Items.Add("Cell 16");
                                    comboBox3.Items.Add("Cell 17");
                                    comboBox3.Items.Add("Cell 18");
                                    comboBox3.Items.Add("Cell 19");
                                    comboBox3.Items.Add("Cell 20");
                                    comboBox3.Items.Add("Cell 21");
                                    break;
                                case "10":
                                    // Battery combobox
                                    comboBox2.Items.Clear();
                                    comboBox2.Items.Add("Voltage 1");
                                    comboBox2.Items.Add("Voltage 2");
                                    comboBox2.Items.Add("Voltage 3");
                                    comboBox2.Items.Add("Voltage 4");
                                    comboBox2.Items.Add("Current");
                                    comboBox2.Items.Add("Temperature 1");
                                    comboBox2.Items.Add("Temperature 2");
                                    comboBox2.Items.Add("Temperature 3");
                                    comboBox2.Items.Add("Temperature 4");
                                    // Cells combobox
                                    comboBox3.Items.Clear();
                                    comboBox3.Items.Add("Current Voltages");
                                    break;
                                case "21":
                                    // Battery combobox
                                    comboBox2.Items.Clear();
                                    comboBox2.Items.Add("Voltage");
                                    comboBox2.Items.Add("Current");
                                    comboBox2.Items.Add("Temperature 1");
                                    comboBox2.Items.Add("Temperature 2");
                                    comboBox2.Items.Add("Temperature 3");
                                    comboBox2.Items.Add("Temperature 4");
                                    // Cells combobox
                                    comboBox3.Items.Clear();
                                    comboBox3.Items.Add("Cell Voltages");
                                    comboBox3.Items.Add("Cell 1");
                                    comboBox3.Items.Add("Cell 2");
                                    comboBox3.Items.Add("Cell 3");
                                    comboBox3.Items.Add("Cell 4");
                                    comboBox3.Items.Add("Cell 5");
                                    comboBox3.Items.Add("Cell 6");
                                    comboBox3.Items.Add("Cell 7");
                                    comboBox3.Items.Add("Cell 8");
                                    comboBox3.Items.Add("Cell 9");
                                    comboBox3.Items.Add("Cell 10");
                                    comboBox3.Items.Add("Cell 11");
                                    comboBox3.Items.Add("Cell 12");
                                    comboBox3.Items.Add("Cell 13");
                                    comboBox3.Items.Add("Cell 14");
                                    comboBox3.Items.Add("Cell 15");
                                    comboBox3.Items.Add("Cell 16");
                                    comboBox3.Items.Add("Cell 17");
                                    comboBox3.Items.Add("Cell 18");
                                    comboBox3.Items.Add("Cell 19");
                                    comboBox3.Items.Add("Cell 20");
                                    comboBox3.Items.Add("Cell 21");
                                    break;
                                default:
                                    // Battery combobox
                                    comboBox2.Items.Clear();
                                    comboBox2.Items.Add("Voltage 1");
                                    comboBox2.Items.Add("Voltage 2");
                                    comboBox2.Items.Add("Voltage 3");
                                    comboBox2.Items.Add("Voltage 4");
                                    comboBox2.Items.Add("Current");
                                    comboBox2.Items.Add("Temperature 1");
                                    comboBox2.Items.Add("Temperature 2");
                                    comboBox2.Items.Add("Temperature 3");
                                    comboBox2.Items.Add("Temperature 4");
                                    // Cells combobox
                                    comboBox3.Items.Clear();
                                    comboBox3.Items.Add("Cell Voltages");
                                    comboBox3.Items.Add("Cell 1");
                                    comboBox3.Items.Add("Cell 2");
                                    comboBox3.Items.Add("Cell 3");
                                    comboBox3.Items.Add("Cell 4");
                                    comboBox3.Items.Add("Cell 5");
                                    comboBox3.Items.Add("Cell 6");
                                    comboBox3.Items.Add("Cell 7");
                                    comboBox3.Items.Add("Cell 8");
                                    comboBox3.Items.Add("Cell 9");
                                    comboBox3.Items.Add("Cell 10");
                                    comboBox3.Items.Add("Cell 11");
                                    comboBox3.Items.Add("Cell 12");
                                    comboBox3.Items.Add("Cell 13");
                                    comboBox3.Items.Add("Cell 14");
                                    comboBox3.Items.Add("Cell 15");
                                    comboBox3.Items.Add("Cell 16");
                                    comboBox3.Items.Add("Cell 17");
                                    comboBox3.Items.Add("Cell 18");
                                    comboBox3.Items.Add("Cell 19");
                                    comboBox3.Items.Add("Cell 20");
                                    comboBox3.Items.Add("Cell 21");
                                    comboBox3.Items.Add("Cell 22");
                                    comboBox3.Items.Add("Cell 23");
                                    comboBox3.Items.Add("Cell 24");
                                    break;
                            }// end switch

                            radioButton1.Enabled = true;
                            radioButton2.Enabled = true;
                            // load saved values here!
                            updateR1((bool)gs.Rows[currentRow][0]);
                            updateR2(!(bool)gs.Rows[currentRow][0]);
                        });// end invoke

                        //take a break
                        Thread.Sleep(100);

                        this.Invoke((MethodInvoker)delegate()
                        {
                            if (radioButton1.Checked == true)
                            {

                                updateC2(gs.Rows[currentRow][1].ToString());
                                comboBox3.SelectedIndex = 0;
                                if (comboBox2.Text == "") { comboBox2.SelectedIndex = 0; }
                            }
                            else
                            {
                                comboBox2.SelectedIndex = 0;
                                updateC3(gs.Rows[currentRow][1].ToString());
                                if (comboBox3.Text == "") { comboBox3.SelectedIndex = 0; }
                            }
                            comboBox2.Enabled = true;
                            comboBox3.Enabled = true;

                            if (cellCable == "10") { radioButton2.Text = "Cur Vs"; }
                            else { radioButton2.Text = "Cells"; }
                            lockUpdate = false;

                        });// end invoke
                    }// end else
                });// end helper thread
            

        }// end function

        // this will update the gs datatable when the radio buttons are changed
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            ThreadPool.QueueUserWorkItem(s =>
               {
                   this.Invoke((MethodInvoker)delegate()
                   {
                       if (comboBox3.Text != "") { gs.Rows[dataGridView1.CurrentRow.Index][0] = radioButton1.Checked; }
                   });
               });
        }

        // this will update the gs datatable when the comboboxes are changed
        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            ThreadPool.QueueUserWorkItem(s =>
               {
                   try
                   {
                       if (radioButton1.Checked == true)
                       {
                           this.Invoke((MethodInvoker)delegate()
                           {
                               if (comboBox3.Text != "") { gs.Rows[dataGridView1.CurrentRow.Index][1] = comboBox2.Text; }
                           });
                       }
                   }
                   catch { }
               });

        }
        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            ThreadPool.QueueUserWorkItem(s =>
               {
                   try
                   {

                       if (radioButton1.Checked == false)
                       {
                           this.Invoke((MethodInvoker)delegate()
                           {
                               if (comboBox3.Text != "") { gs.Rows[dataGridView1.CurrentRow.Index][1] = comboBox3.Text; }
                           });
                       }
                   }
                   catch { }
               });
        }

        //////////////////////////////////////////////////////////////////////locking stuff////////////////////////////
        private readonly object combo2Lock = new object();

        private void updateC2(string inVal)
        {
            lock (combo2Lock)
            {
                comboBox2.Text = inVal;
            }
        }

        private readonly object combo3Lock = new object();

        private void updateC3(string inVal)
        {
            lock (combo3Lock)
            {
                comboBox3.Text = inVal;
            }
        }

        private readonly object radio1Lock = new object();

        private void updateR1(bool inVal)
        {
            lock (radio1Lock)
            {
                radioButton1.Checked = inVal;
            }
        }

        private readonly object radio2Lock = new object();

        private void updateR2(bool inVal)
        {
            lock (radio2Lock)
            {
                radioButton2.Checked = inVal;
            }
        }

    }
}
