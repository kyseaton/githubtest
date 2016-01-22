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

                                    int chargerNum = -1;
                                    //figure out which charger you have on that channel.
                                    if (d.Rows[currentRow][9].ToString().Length == 0)
                                    {
                                        //no id
                                    }
                                    else if (d.Rows[currentRow][9].ToString().Length < 3)
                                    {
                                        chargerNum = int.Parse(d.Rows[currentRow][9].ToString());
                                    }
                                    else if (d.Rows[currentRow][9].ToString().Length == 3)
                                    {
                                        chargerNum = int.Parse(d.Rows[currentRow][9].ToString().Substring(0, 1));
                                    }
                                    else if (d.Rows[currentRow][9].ToString().Length == 4)
                                    {
                                        chargerNum = int.Parse(d.Rows[currentRow][9].ToString().Substring(0, 2));
                                    }

                                    if ((bool)d.Rows[currentRow][4] && tempClick == currentRow)  // test to see if we've clicked in the mean time...
                                    {

                                        //put this new data in the chart!
                                        this.Invoke((MethodInvoker)delegate
                                        {

                                            // first set the cell to green
                                            dataGridView1.Rows[currentRow].Cells[4].Style.BackColor = Color.Green;

                                            //update chart function
                                            updateChart(testData);

                                            // check if we are a slave row..
                                            int masterRow = -1;
                                            if (d.Rows[currentRow][9].ToString().Length > 2 && d.Rows[currentRow][9].ToString().Contains("S"))
                                            {
                                                // we have a slave
                                                string temp = d.Rows[currentRow][9].ToString().Replace("-S", "");
                                                for (int q = 0; q < 16; q++)
                                                {
                                                    if (d.Rows[q][9].ToString().Contains(temp) && d.Rows[q][9].ToString().Contains("M"))
                                                    {
                                                        masterRow = q;
                                                        break;
                                                    }
                                                }
                                            }



                                            //Real Time Data Portion
                                            string tempText = "";
                                            tempText = System.DateTime.Now.ToString("M/d/yyyy") + "       Terminal:  " + testData.terminalID.ToString() + Environment.NewLine;
                                            tempText += "Temp. Cable:  " + (3 - (masterRow != -1 ? GlobalVars.CScanData[masterRow].TCAB : testData.TCAB)).ToString() + "   (" + (masterRow != -1 ? GlobalVars.CScanData[masterRow].tempPlateType : testData.tempPlateType) + ")" + Environment.NewLine;
                                            tempText += "Cells Cable:  " + testData.CCID.ToString() + "   (" + testData.cellCableType + ")" + Environment.NewLine;
                                            tempText += "Shunt Cable:  " + (masterRow != -1 ? GlobalVars.CScanData[masterRow].SHCID : testData.SHCID).ToString() + "   (" + (masterRow != -1 ? GlobalVars.CScanData[masterRow].shuntCableType : testData.shuntCableType) + ")" + Environment.NewLine;
                                            tempText += Environment.NewLine;
                                            if (testData.cellCableType != "NONE")       // don't show these when we have no cable attached...
                                            {
                                                if (testData.cellCableType == "4 BATT" || testData.cellCableType == "CELL SIM" || testData.cellCableType == "2X11 Cable" || testData.cellCableType == "3X7 Cable" || testData.cellCableType == "Unknown Cable")
                                                {
                                                    tempText += "Voltage Batt 1:  " + testData.VB1.ToString("00.00") + Environment.NewLine;
                                                }
                                                else
                                                {
                                                    tempText += "Voltage Batt:  " + testData.VB1.ToString("00.00") + Environment.NewLine;
                                                }
                                                if (testData.cellCableType == "4 BATT" || testData.cellCableType == "CELL SIM" || testData.cellCableType == "2X11 Cable" || testData.cellCableType == "3X7 Cable" || testData.cellCableType == "Unknown Cable")
                                                {
                                                    tempText += "Voltage Batt 2:  " + testData.VB2.ToString("00.00") + Environment.NewLine;
                                                    if (testData.cellCableType == "4 BATT" || testData.cellCableType == "CELL SIM" || testData.cellCableType == "3X7 Cable" || testData.cellCableType == "Unknown Cable")
                                                    {
                                                        tempText += "Voltage Batt 3:  " + testData.VB3.ToString("00.00") + Environment.NewLine;
                                                        if (testData.cellCableType == "4 BATT" || testData.cellCableType == "CELL SIM" || testData.cellCableType == "Unknown Cable")
                                                        {
                                                            tempText += "Voltage Batt 4:  " + testData.VB4.ToString("00.00") + Environment.NewLine;
                                                        }
                                                    }
                                                }
                                            }
                                            tempText += Environment.NewLine;
                                            // select which currents to display

                                            // check if we need to use the master Current..
                                            
                                            if (masterRow != -1)
                                            {
                                                if (GlobalVars.CScanData[masterRow].shuntCableType == "NONE")
                                                {
                                                    //skip the current
                                                }
                                                else if (GlobalVars.CScanData[masterRow].shuntCableType == "TEST BOX")
                                                {
                                                    //dispaly both ...
                                                    tempText += "Current#1:  " + GlobalVars.CScanData[masterRow].currentOne.ToString("00.00") + Environment.NewLine;
                                                    tempText += "Current#2:  " + GlobalVars.CScanData[masterRow].currentTwo.ToString("00.00") + Environment.NewLine;
                                                }
                                                //if we have a mini that is charging...
                                                else if (d.Rows[currentRow][10].ToString().Contains("mini") && !(GlobalVars.ICData[chargerNum].testMode.ToString().Contains("Cap") || GlobalVars.ICData[chargerNum].testMode.ToString().Contains("Discharge")))
                                                {
                                                    // for the mini case
                                                    tempText += "Current:  " + GlobalVars.CScanData[masterRow].currentTwo.ToString("00.000") + Environment.NewLine;
                                                }
                                                else if (d.Rows[currentRow][10].ToString().Contains("mini"))
                                                {
                                                    // all other cases
                                                    tempText += "Current:  " + (GlobalVars.CScanData[masterRow].currentOne).ToString("00.000") + Environment.NewLine;
                                                }
                                                else if (GlobalVars.CScanData[masterRow].shuntCableType == "100A")
                                                {
                                                    // all other cases
                                                    tempText += "Current:  " + GlobalVars.CScanData[masterRow].currentOne.ToString("00.0") + Environment.NewLine;
                                                }
                                                else
                                                {
                                                    // all other cases
                                                    tempText += "Current:  " + GlobalVars.CScanData[masterRow].currentOne.ToString("00.00") + Environment.NewLine;
                                                }
                                            }
                                            // normal case
                                            else
                                            {
                                                if (testData.shuntCableType == "NONE")
                                                {
                                                    //skip it...
                                                }
                                                else if (testData.shuntCableType == "TEST BOX")
                                                {
                                                    //dispaly both ...
                                                    tempText += "Current#1:  " + testData.currentOne.ToString("00.00") + Environment.NewLine;
                                                    tempText += "Current#2:  " + testData.currentTwo.ToString("00.00") + Environment.NewLine;
                                                }
                                                //if we have a mini that is charging...
                                                else if (d.Rows[currentRow][10].ToString().Contains("mini") && !(GlobalVars.ICData[chargerNum].testMode.ToString().Contains("Cap") || GlobalVars.ICData[chargerNum].testMode.ToString().Contains("Discharge")))
                                                {
                                                    // for the mini case
                                                    tempText += "Current:  " + testData.currentTwo.ToString("00.000") + Environment.NewLine;
                                                }
                                                else if (d.Rows[currentRow][10].ToString().Contains("mini"))
                                                {
                                                    // all other cases
                                                    tempText += "Current:  " + (testData.currentOne).ToString("00.000") + Environment.NewLine;
                                                }
                                                else if (testData.shuntCableType == "100A")
                                                {
                                                    // all other cases
                                                    tempText += "Current:  " + testData.currentOne.ToString("00.0") + Environment.NewLine;
                                                }
                                                else
                                                {
                                                    // all other cases
                                                    tempText += "Current:  " + testData.currentOne.ToString("00.00") + Environment.NewLine;
                                                }
                                            }

                                            tempText += Environment.NewLine;

                                            int cellsToDisplay = 0;
                                            if ((int) pci.Rows[currentRow][3] != -1 && (int) pci.Rows[currentRow][3] <= GlobalVars.CScanData[currentRow].cellsToDisplay)
                                            {
                                                cellsToDisplay = (int)pci.Rows[currentRow][3];
                                            }
                                            else
                                            {
                                                cellsToDisplay = GlobalVars.CScanData[currentRow].cellsToDisplay;
                                            }
                                            if (GlobalVars.Pos2Neg == false)
                                            {
                                                for (int i = 0; i < cellsToDisplay; i++)
                                                {
                                                    tempText += "Cell #" + (i + 1).ToString() + ":  " + testData.orderedCells[i].ToString("0.000") + Environment.NewLine;
                                                }
                                            }
                                            else
                                            {
                                                for (int i = 0; i < cellsToDisplay; i++)
                                                {
                                                    tempText += "Cell #" + (i + 1).ToString() + ":  " + testData.orderedCells[cellsToDisplay - i - 1].ToString("0.000") + Environment.NewLine;
                                                }
                                            }

                                            
                                            tempText += Environment.NewLine;

                                            if ((masterRow != -1 ?  GlobalVars.CScanData[masterRow].tempPlateType != "NONE" : testData.tempPlateType != "NONE"))
                                            { 
                                                // WE need to display open when we get -99, cold for -98, hot for -97 and shorted for -96
                                                switch (masterRow != -1 ? Convert.ToInt16(GlobalVars.CScanData[masterRow].TP1) : Convert.ToInt16(testData.TP1))
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
                                                        tempText += "Temp Plate 1:  " + (GlobalVars.useF ? ConvertCelsiusToFahrenheit(masterRow != -1 ? GlobalVars.CScanData[masterRow].TP1 : testData.TP1).ToString("00.00") : (masterRow != -1 ? GlobalVars.CScanData[masterRow].TP1 : testData.TP1).ToString("00.0")) + Environment.NewLine;
                                                        break;
                                                }
                                                switch (masterRow != -1 ? Convert.ToInt16(GlobalVars.CScanData[masterRow].TP2) : Convert.ToInt16(testData.TP2))
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
                                                        tempText += "Temp Plate 2:  " + (GlobalVars.useF ? ConvertCelsiusToFahrenheit(masterRow != -1 ? GlobalVars.CScanData[masterRow].TP2 : testData.TP2).ToString("00.00") : (masterRow != -1 ? GlobalVars.CScanData[masterRow].TP2 : testData.TP2).ToString("00.0")) + Environment.NewLine;
                                                        break;
                                                }
                                                switch (masterRow != -1 ? Convert.ToInt16(GlobalVars.CScanData[masterRow].TP3) : Convert.ToInt16(testData.TP3))
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
                                                        tempText += "Temp Plate 3:  " + (GlobalVars.useF ? ConvertCelsiusToFahrenheit(masterRow != -1 ? GlobalVars.CScanData[masterRow].TP3 : testData.TP3).ToString("00.00") : (masterRow != -1 ? GlobalVars.CScanData[masterRow].TP3 : testData.TP3).ToString("00.0")) + Environment.NewLine;
                                                        break;
                                                }
                                                switch (masterRow != -1 ? Convert.ToInt16(GlobalVars.CScanData[masterRow].TP4) : Convert.ToInt16(testData.TP4))
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
                                                        tempText += "Temp Plate 4:  " + (GlobalVars.useF ? ConvertCelsiusToFahrenheit(masterRow != -1 ? GlobalVars.CScanData[masterRow].TP4 : testData.TP4).ToString("00.00") : (masterRow != -1 ? GlobalVars.CScanData[masterRow].TP4 : testData.TP4).ToString("00.0")) + Environment.NewLine;
                                                        break;
                                                }
                                                switch (masterRow != -1 ? Convert.ToInt16(GlobalVars.CScanData[masterRow].TP5) : Convert.ToInt16(testData.TP5))
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
                                                        tempText += "Ambient Temp:  " + (GlobalVars.useF ? ConvertCelsiusToFahrenheit(masterRow != -1 ? GlobalVars.CScanData[masterRow].TP5 : testData.TP5).ToString("00.00") : (masterRow != -1 ? GlobalVars.CScanData[masterRow].TP5 : testData.TP5).ToString("00.0")) + Environment.NewLine;
                                                        break;
                                                }
                                            }

                                            tempText += Environment.NewLine;

                                            tempText += "Reference Voltage:  " + testData.ref95V.ToString("0.000") + Environment.NewLine;

                                            tempText += Environment.NewLine;

                                            tempText += "CSCAN Program Version: " + testData.programVersion + Environment.NewLine;
                                            if (d.Rows[currentRow][10].ToString().Contains("ICA"))
                                            {
                                                tempText += "IC Program Version: " + GlobalVars.ICData[chargerNum].PV1.ToString() + Environment.NewLine;
                                                tempText += "IC COMS Program Version:  " + GlobalVars.ICData[chargerNum].PV2.ToString() + "";
                                            }

                                            LockWindowUpdate(label1.Handle);
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
                                                if (d.Rows[q][9].ToString().Contains(temp) && d.Rows[q][9].ToString().Contains("S"))
                                                {
                                                    slaveRow = q;
                                                    break;
                                                }
                                            }
                                        }

                                        if ((bool)d.Rows[currentRow][4] &&
                                            //(bool)d.Rows[currentRow][8] &&
                                            !d.Rows[currentRow][9].ToString().Contains("S") &&
                                            GlobalVars.CScanData[currentRow].connected &&
                                            !d.Rows[currentRow][10].ToString().Contains("ICA") &&
                                            (d.Rows[currentRow][10].ToString() == "" || dataGridView1.Rows[currentRow].Cells[8].Style.BackColor != Color.Olive || dataGridView1.Rows[currentRow].Cells[8].Style.BackColor != Color.Red))  // if a charger type isn't already there maybe we need to update with a CSCAN controlled charger...
                                        {
                                            // we got a CSCAN connected charger...
                                            updateD(currentRow, 10, (GlobalVars.CScanData[currentRow].cellCableType == "CELL SIM" ? "CCA (SIM)" : "CCA"));
                                            if (slaveRow > -1)
                                            {
                                                updateD(slaveRow, 10, (GlobalVars.CScanData[currentRow].cellCableType == "CELL SIM" ? "CCA (SIM)" : "CCA"));
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
                                                if ((bool)d.Rows[currentRow][5] == false)
                                                {
                                                    updateD(currentRow, 11, "Power Off");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 11, "Power Off");
                                                    }
                                                }
                                                else
                                                {
                                                    updateD(currentRow, 11, "Power Fail");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 11, "Power Fail");
                                                    }
                                                }
                                            }
                                        }
                                        //
                                        else if (
                                            //(bool)d.Rows[currentRow][8] && 
                                            GlobalVars.CScanData[currentRow].connected == false && 
                                            !d.Rows[currentRow][9].ToString().Contains("S") && 
                                            d.Rows[currentRow][10].ToString().Contains("CCA"))
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
                                        else if (dataGridView1.Rows[currentRow].Cells[8].Style.BackColor == Color.Olive && GlobalVars.CScanData[currentRow].powerOn == false && !d.Rows[currentRow][9].ToString().Contains("S") && d.Rows[currentRow][10].ToString().Contains("CCA"))
                                        {
                                            this.Invoke((MethodInvoker)delegate 
                                            { 
                                                dataGridView1.Rows[currentRow].Cells[8].Style.BackColor = Color.Red;
                                                if (slaveRow > -1)
                                                {
                                                    dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.Red;
                                                }
                                            });
                                            if ((bool)d.Rows[currentRow][5] == false)
                                            {
                                                updateD(currentRow, 11, "Power Off");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 11, "Power Off");
                                                }
                                            }
                                            else
                                            {
                                                updateD(currentRow, 11, "Power Fail");
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 11, "Power Fail");
                                                }
                                            }
                                        }
                                        else if (dataGridView1.Rows[currentRow].Cells[8].Style.BackColor == Color.Red && GlobalVars.CScanData[currentRow].powerOn && !d.Rows[currentRow][9].ToString().Contains("S") && d.Rows[currentRow][10].ToString().Contains("CCA"))
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

                                        if ((bool)d.Rows[currentRow][4] &&
                                            //(bool)d.Rows[currentRow][8] && 
                                            d.Rows[currentRow][10].ToString() == "" && 
                                            !d.Rows[currentRow][9].ToString().Contains("S") && 
                                            GlobalVars.CScanData[currentRow].shuntCableType != "NONE")
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
                                        else if ((bool)d.Rows[currentRow][4] && d.Rows[currentRow][10].ToString() == "Shunt" && !d.Rows[currentRow][9].ToString().Contains("S") && GlobalVars.CScanData[currentRow].shuntCableType != "NONE")
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
                                        else if ((bool)d.Rows[currentRow][4] && d.Rows[currentRow][10].ToString() == "Shunt" && !d.Rows[currentRow][9].ToString().Contains("S") && GlobalVars.CScanData[currentRow].shuntCableType == "NONE")
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
                                                if ((bool)d.Rows[currentRow][5] == false)
                                                {
                                                    updateD(currentRow, 11, "Power Off");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 11, "Power Off");
                                                    }
                                                }
                                                else
                                                {
                                                    updateD(currentRow, 11, "Power Fail");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 11, "Power Fail");
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                updateD(currentRow, 11, (GlobalVars.cHold[currentRow] ? "HOLD" : ((bool) d.Rows[currentRow][8] ? "RUN" : "Not Controlled")));
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 11, (GlobalVars.cHold[currentRow] ? "HOLD" : ((bool)d.Rows[currentRow][8] ? "RUN" : "Not Controlled")));
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

                                        return;
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
                                    chart1.Titles.Clear();
                                });

                            }

                            /// NON Selected Row Case////////////////////////////////////////////////////////////////////////
                            // now look at all of the other cases to up date the label after a little break...
                            // if we are not looking for stations with the "find stations" function...

                            if (toolStripMenuItem34.Enabled == true)
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
                                                if (d.Rows[q][9].ToString().Contains(temp) &&
                                                    d.Rows[q][9].ToString().Contains("S") //&&
                                                    //(bool) d.Rows[q][8]
                                                    )
                                                {
                                                    slaveRow = q;
                                                    break;
                                                }
                                            }
                                        }

                                        if ((bool)d.Rows[j][4] && !d.Rows[j][9].ToString().Contains("S"))
                                        {
                                            if ((bool)d.Rows[j][4] &&
                                                //(bool)d.Rows[j][8] &&
                                                GlobalVars.CScanData[j].connected &&
                                                !d.Rows[j][10].ToString().Contains("ICA") &&
                                                (d.Rows[j][10].ToString() == "" || dataGridView1.Rows[j].Cells[8].Style.BackColor != Color.Olive || dataGridView1.Rows[j].Cells[8].Style.BackColor != Color.Red))  // if a charger type isn't already there maybe we need to update with a CSCAN controlled charger...
                                            {
                                                // we got a CSCAN connected charger...
                                                updateD(j, 10, (GlobalVars.CScanData[j].cellCableType == "CELL SIM" ? "CCA (SIM)" : "CCA"));
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 10, (GlobalVars.CScanData[j].cellCableType == "CELL SIM" ? "CCA (SIM)" : "CCA"));
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
                                                    if ((bool)d.Rows[j][5] == false)
                                                    {
                                                        updateD(j, 11, "Power Off");
                                                        if (slaveRow > -1)
                                                        {
                                                            updateD(slaveRow, 11, "Power Off");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        updateD(j, 11, "Power Fail");
                                                        if (slaveRow > -1)
                                                        {
                                                            updateD(slaveRow, 11, "Power Fail");
                                                        }
                                                    }
                                                }
                                            }
                                            //
                                            else if (
                                                //(bool)d.Rows[j][8] && 
                                                GlobalVars.CScanData[j].connected == false && 
                                                d.Rows[j][10].ToString().Contains("CCA") && 
                                                !d.Rows[j][9].ToString().Contains("S"))
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
                                            else if (dataGridView1.Rows[j].Cells[8].Style.BackColor == Color.Olive && GlobalVars.CScanData[j].powerOn == false && d.Rows[j][10].ToString().Contains("CCA") && !d.Rows[j][9].ToString().Contains("S"))
                                            {
                                                this.Invoke((MethodInvoker)delegate 
                                                { 
                                                    dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Red;
                                                    if (slaveRow > -1)
                                                    {
                                                        dataGridView1.Rows[slaveRow].Cells[8].Style.BackColor = Color.Red;
                                                    }
                                                });

                                                if ((bool)d.Rows[j][5] == false)
                                                {
                                                    updateD(j, 11, "Power Off");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 11, "Power Off");
                                                    }
                                                }
                                                else
                                                {
                                                    updateD(j, 11, "Power Fail");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 11, "Power Fail");
                                                    }
                                                }
                                            }
                                            else if (dataGridView1.Rows[j].Cells[8].Style.BackColor == Color.Red && GlobalVars.CScanData[j].powerOn && d.Rows[j][10].ToString().Contains("CCA") && !d.Rows[j][9].ToString().Contains("S"))
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

                                        if ((bool)d.Rows[j][4] &&
                                            //(bool)d.Rows[j][8] && 
                                            d.Rows[j][10].ToString() == "" && 
                                            GlobalVars.CScanData[j].shuntCableType != "NONE" && 
                                            !d.Rows[j][9].ToString().Contains("S"))
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
                                        else if ((bool)d.Rows[j][4] && d.Rows[j][10].ToString() == "Shunt" && GlobalVars.CScanData[j].shuntCableType != "NONE" && !d.Rows[j][9].ToString().Contains("S"))
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
                                        else if ((bool)d.Rows[j][4] && 
                                            //(bool)d.Rows[j][8] && 
                                            d.Rows[j][10].ToString() == "Shunt" && 
                                            GlobalVars.CScanData[j].shuntCableType == "NONE" && 
                                            !d.Rows[j][9].ToString().Contains("S"))
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
                                                if ((bool)d.Rows[j][5] == false)
                                                {
                                                    updateD(j, 11, "Power Off");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 11, "Power Off");
                                                    }
                                                }
                                                else
                                                {
                                                    updateD(j, 11, "Power Fail");
                                                    if (slaveRow > -1)
                                                    {
                                                        updateD(slaveRow, 11, "Power Fail");
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                updateD(j, 11, (GlobalVars.cHold[j] ? "HOLD" : ((bool)d.Rows[j][8] ? "RUN" : "Not Controlled")));
                                                if (slaveRow > -1)
                                                {
                                                    updateD(slaveRow, 11, (GlobalVars.cHold[j] ? "HOLD" : ((bool)d.Rows[j][8] ? "RUN" : "Not Controlled")));
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

                                            return;
                                        }

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
                            //MessageBox.Show(ex.ToString());
                        }
                        
                    }                   // end catch
                }                       // end while (this is an endless loop, only the cancel token kills it)
            }, cPollCScans.Token);                     // end thread

        }

        private void updateChart(CScanDataStore testData)
        {
            try
            {
                // first check to see if you have a cells cable...
                if (testData.cellCableType == "NONE")
                {
                    //clear and return
                    chart1.Series.Clear();
                    chart1.Titles.Clear();
                    return;
                }

                //Replace based on values selected in radio1("Battery") or radio2 ("Cells")
                //and combo2 (Battery voltages) or combo3 (Cell voltages)
                //if that row is selected, update the chart portion
                int station = dataGridView1.CurrentRow.Index;

                int Cells;

                if ((int)pci.Rows[station][3] == -1)
                {
                    Cells = GlobalVars.CScanData[station].cellsToDisplay;
                }
                else
                {
                    Cells = (int)pci.Rows[station][3];
                }

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

                    series1.Points.AddXY(1, testData.VB1);
                    series1.Points[0].Color = pointColorMain(station, testData.VB1, false);
                    series1.Points[0].Label = "VB1";
                    series1.Points.AddXY(2, testData.VB2);
                    series1.Points[1].Color = pointColorMain(station, testData.VB2, false);
                    series1.Points[1].Label = "VB2";
                    series1.Points.AddXY(3, testData.VB3);
                    series1.Points[2].Color = pointColorMain(station, testData.VB3, false);
                    series1.Points[2].Label = "VB3";
                    series1.Points.AddXY(4, testData.VB4);
                    series1.Points[3].Color = pointColorMain(station, testData.VB4, false);
                    series1.Points[3].Label = "VB4";

                    chart1.Titles.Clear();
                    chart1.Titles.Add("Current Voltages");
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

                    //special cable cases first...
                    if (GlobalVars.CScanData[station].CCID == 3)
                    {
                        // this is the 2X 11 cable
                        for (int i = 0; i < Cells; i++)
                        {
                            if (GlobalVars.Pos2Neg == false)
                            {
                                series1.Points.AddXY(i % 11 + 1, testData.orderedCells[i]);
                                // color test
                                series1.Points[i].Color = pointColorMain(station, testData.orderedCells[i], true);
                                if (i == 10)
                                {
                                    // add a blank point
                                    series1.Points.AddXY(" ", 0);
                                }
                            }
                            else
                            {
                                series1.Points.AddXY(i % 11 + 1, testData.orderedCells[Cells - i - 1]);
                                // color test
                                series1.Points[i].Color = pointColorMain(station, testData.orderedCells[Cells - i - 1], true);
                                if (i == 6 || i == 13)
                                {
                                    // add a blank point
                                    series1.Points.AddXY(" ", 0);
                                }
                            }
                        }
                    }
                    else if (GlobalVars.CScanData[station].CCID == 4)
                    {
                        // this is the 3X7 cable
                        for (int i = 0; i < Cells; i++)
                        {
                            if (GlobalVars.Pos2Neg == false)
                            {
                                series1.Points.AddXY(i % 7 + 1, testData.orderedCells[i]);
                                // color test
                                series1.Points[i].Color = pointColorMain(station, testData.orderedCells[i], true);
                                if (i == 10)
                                {
                                    // add a blank point
                                    series1.Points.AddXY(" ", 0);
                                }
                            }
                            else
                            {
                                series1.Points.AddXY(i % 7 + 1, testData.orderedCells[Cells - i - 1]);
                                // color test
                                series1.Points[i].Color = pointColorMain(station, testData.orderedCells[Cells - i - 1], true);
                                if (i == 6 || i == 13)
                                {
                                    // add a blank point
                                    series1.Points.AddXY(" ", 0);
                                }
                            }
                        }
                    }
                    else
                    {
                        //  all other cases...
                        for (int i = 0; i < Cells; i++)
                        {
                            if (GlobalVars.Pos2Neg == false)
                            {
                                series1.Points.AddXY(i + 1, testData.orderedCells[i]);
                                // color test
                                series1.Points[i].Color = pointColorMain(station, testData.orderedCells[i], true);
                            }
                            else
                            {
                                series1.Points.AddXY(i + 1, testData.orderedCells[Cells - i - 1]);
                                // color test
                                series1.Points[i].Color = pointColorMain(station, testData.orderedCells[Cells - i - 1], true);
                            }
                        }
                    }
                    chart1.Titles.Clear();
                    if (Cells != 0) { chart1.Titles.Add("Cell Voltages"); }
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
                        if (comboBox2.SelectedIndex < 0) 
                        {
                            //reset to the default..
                            radioButton2.Checked = true;
                            comboBox3.SelectedIndex = 0;
                            if (testData.CCID == 10)
                            {
                                comboBox3.Text = "Current Votages";
                            }
                            else
                            {
                                comboBox3.Text = "Cell Voltages";
                            }
                            return; 
                        }
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
                                //if we have a mini that is charging...
                                if (d.Rows[testData.terminalID][10].ToString().Contains("mini") && !(d.Rows[testData.terminalID][2].ToString().Contains("Cap") || d.Rows[testData.terminalID][2].ToString().Contains("Discharge")))
                                {
                                    q = 9;
                                }
                                else
                                {
                                    q = 8;
                                }
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
                            series1.Points.AddXY(Math.Round(double.Parse(graphMainSet.Tables[0].Rows[i][7].ToString()) * 1440), float.Parse(graphMainSet.Tables[0].Rows[i][q].ToString()));
                            // color test
                            if (q == 8 || q == 7)
                            {
                                //current
                                series1.Points[i].Color = Color.Blue;
                            }
                            else if (q > 37 && q < 42)
                            {
                                //temp
                                series1.Points[i].Color = Color.LightSeaGreen;
                            }
                            else
                            {
                                series1.Points[i].Color = pointColorMain(station, double.Parse(graphMainSet.Tables[0].Rows[i][q].ToString()), false);
                            }
                        }

                        // pad with zero Vals to help with the look of the plot...
                        // first get the interval and total points
                        int interval = 1;
                        int points = 1;

                        switch (d.Rows[station][2].ToString())
                        {
                            case "As Received":
                                interval = 1 / 30;
                                points = 3;
                                break;
                            case "Full Charge-6":
                            case "Combo: >>FC-6<<  Cap-1":
                            case "Combo: FC-6 >><< Cap-1":
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
                            case "Combo: FC-6  >>Cap-1<<":
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


                        if (graphMainSet.Tables[0].Rows.Count <= points - 1)
                        {
                            for (int i = graphMainSet.Tables[0].Rows.Count; i <= points - 1; i++)
                            {
                                series1.Points.AddXY(i * interval, 0);
                            }
                        }

                        chart1.Titles.Clear();
                        chart1.Titles.Add(comboBox2.Text);
                        chart1.Invalidate();
                        chart1.ChartAreas[0].RecalculateAxesScale();
                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                    }
                }// end else  if

                    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                // Cells Section!!!!
                else
                {
                    try
                    {

                        // we need to get the number of cells incase we are reversing...
                        int numCells;

                        if ((int)pci.Rows[station][3] == -1)
                        {
                            numCells = GlobalVars.CScanData[station].cellsToDisplay;
                        }
                        else
                        {
                            numCells = (int)pci.Rows[station][3];
                        }

                        chart1.ChartAreas[0].AxisY.Title = "Voltage";
                        chart1.ChartAreas[0].AxisX.Title = "Time";
                        int q;
                        // only do something if the radio button is selected
                        if (radioButton2.Checked == false || comboBox3.SelectedIndex < 0) 
                        {
                            //reset to the default..
                            radioButton2.Checked = true;
                            comboBox3.SelectedIndex = 0;
                            if (testData.CCID == 10)
                            {
                                comboBox3.Text = "Current Votages";
                            }
                            else
                            {
                                comboBox3.Text = "Cell Voltages";
                            }
                            return; 
                        }
                        // Here we will look at the Value selected and then plot graph1Set

                        //find out which graph to plot from the selected text
                        switch (comboBox3.Text)
                        {
                            case "Ending Voltages":
                                q = 999;
                                break;
                            case "Cell 1":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells; }
                                else { q = 14; }
                                break;
                            case "Cell 2":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 1; }
                                else { q = 15; }
                                break;
                            case "Cell 3":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 2; }
                                else { q = 16; }
                                break;
                            case "Cell 4":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 3; }
                                else { q = 17; }
                                break;
                            case "Cell 5":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 4; }
                                else { q = 18; }
                                break;
                            case "Cell 6":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 5; }
                                else { q = 19; }
                                break;
                            case "Cell 7":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 6; }
                                else { q = 20; }
                                break;
                            case "Cell 8":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 7; }
                                else { q = 21; }
                                break;
                            case "Cell 9":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 8; }
                                else { q = 22; }
                                break;
                            case "Cell 10":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 9; }
                                else { q = 23; }
                                break;
                            case "Cell 11":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 10; }
                                else { q = 24; }
                                break;
                            case "Cell 12":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 11; }
                                else { q = 25; }
                                break;
                            case "Cell 13":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 12; }
                                else { q = 26; }
                                break;
                            case "Cell 14":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 13; }
                                else { q = 27; }
                                break;
                            case "Cell 15":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 14; }
                                else { q = 28; }
                                break;
                            case "Cell 16":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 15; }
                                else { q = 29; }
                                break;
                            case "Cell 17":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 16; }
                                else { q = 30; }
                                break;
                            case "Cell 18":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 17; }
                                else { q = 31; }
                                break;
                            case "Cell 19":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 18; }
                                else { q = 32; }
                                break;
                            case "Cell 20":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 19; }
                                else { q = 33; }
                                break;
                            case "Cell 21":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 20; }
                                else { q = 34; }
                                break;
                            case "Cell 22":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 21; }
                                else { q = 35; }
                                break;
                            case "Cell 23":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 22; }
                                else { q = 36; }
                                break;
                            case "Cell 24":
                                if (GlobalVars.Pos2Neg) { q = 13 + numCells - 23; }
                                else { q = 37; }
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
                            for (int i = 0; i < Cells; i++)
                            {
                                series1.Points.AddXY(i + 1, graphMainSet.Tables[0].Rows[graphMainSet.Tables[0].Rows.Count - 1][i + 14]);
                                // color test
                                series1.Points[i].Color = pointColorMain(station, double.Parse(graphMainSet.Tables[0].Rows[graphMainSet.Tables[0].Rows.Count - 1][i + 14].ToString()), true);
                            }
                        }
                        else
                        {
                            for (int i = 0; i < graphMainSet.Tables[0].Rows.Count; i++)
                            {
                                series1.Points.AddXY(Math.Round(double.Parse(graphMainSet.Tables[0].Rows[i][7].ToString()) * 1440), graphMainSet.Tables[0].Rows[i][q]);
                                // color test
                                series1.Points[i].Color = pointColorMain(station, double.Parse(graphMainSet.Tables[0].Rows[i][q].ToString()), true);
                            }
                            // pad with zero Vals to help with the look of the plot...
                            // first get the interval and total points
                            int interval = 1;
                            int points = 1;

                            switch (d.Rows[station][2].ToString())
                            {
                                case "As Received":
                                    interval = 1 / 30;
                                    points = 3;
                                    break;
                                case "Full Charge-6":
                                case "Combo: >>FC-6<<  Cap-1":
                                case "Combo: FC-6 >><< Cap-1":
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
                                case "Combo: FC-6  >>Cap-1<<":
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


                            if (graphMainSet.Tables[0].Rows.Count <= points - 1)
                            {
                                for (int i = graphMainSet.Tables[0].Rows.Count; i <= points - 1; i++)
                                {
                                    series1.Points.AddXY(i * interval, 0);
                                }
                            }
                        }

                        chart1.Titles.Clear();
                        chart1.Titles.Add(comboBox3.Text);
                        chart1.Invalidate();
                        chart1.ChartAreas[0].RecalculateAxesScale();


                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                    }
                }// end else

            }
            catch
            {
                //do nothing...
            }
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
                        if (toolStripMenuItem30.Checked && multi == 0 && toolStripMenuItem34.Enabled == true && !sshold)          // sequential scanning is turned on
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

        private Color pointColorMain(int station, double Value, bool singleCell)
        {

            if (radioButton1.Checked == true)
            {
                switch (comboBox2.Text)
                {
                    case "Current":
                        return System.Drawing.Color.Blue;
                    case "Temperature 1":
                    case "Temperature 2":
                    case "Temperature 3":
                    case "Temperature 4":
                        return System.Drawing.Color.LightSeaGreen;
                    default:
                        break;
                }
            }

            // we use station to look at the particular channel we are working with
            int Cells;
            if (singleCell) 
            {
                Cells = 1; 
            }
            else 
            {
                if ((int)pci.Rows[station][3] == -1)
                {
                    if (GlobalVars.CScanData[station].cellCableType == "4 BATT")
                    {
                        // this simulates a 24V nominal SLA battery...
                        Cells = 20;
                    }
                    else
                    {
                        Cells = GlobalVars.CScanData[station].cellsToDisplay;
                    }
                }
                else
                {
                    Cells = (int)pci.Rows[station][3];
                }
            }

            string tech = pci.Rows[station][1].ToString();
            string test_type = d.Rows[station][2].ToString();

            // test_type is the type of test we are generating the colors for

            // Three types of batteries (NiCd, SLA and NiCd ULM) and two directions (charge discharge)

            // normal vented NiCds
            double Min1 = 0;
            double Min2 = 0;
            double Min3 = 0;
            double Min4 = 0;
            double Max = 0;

            switch (tech)
            {
                case "NiCd":
                    // Discharge
                    if (test_type == "As Received" || test_type == "Capacity-1" || test_type == "Discharge" || test_type == "Custom Cap" || test_type == "Combo: FC-6  >>Cap-1<<" || test_type == "")
                    {
                        Min4 = 1 * Cells;
                        Max = 1.05 * Cells;

                        if (Value > Max) { return Color.Green; }
                        else if (Value > Min4) { return Color.Orange; }
                        else { return Color.Red; }

                    }
                    // Charge
                    else
                    {
                        Min1 = 0.25 * Cells;
                        Min2 = 1.5 * Cells;
                        Min3 = 1.55 * Cells;
                        Max = ((-1 == (float)pci.Rows[station][7]) ? 1.75 : (float)pci.Rows[station][7]) * Cells;

                        if (Value > Max) { return Color.Red; }
                        else if (Value > Min3) { return Color.Green; }
                        else if (Value > Min2) { return Color.Orange; }
                        else if (Value > Min1) { return Color.Yellow; }
                        else { return Color.Red; }
                    }
                case "Sealed Lead Acid":
                    // Discharge
                    if (test_type == "As Received" || test_type == "Capacity-1" || test_type == "Discharge" || test_type == "Custom Cap" || test_type == "")
                    {

                        Min4 = (20.0 / 24) * (float)pci.Rows[station][2];
                        Max = (21.0 /24) * (float) pci.Rows[station][2];

                        if (Value > Max) { return Color.Green; }
                        else if (Value > Min4) { return Color.Orange; }
                        else { return Color.Red; }

                    }
                    // Charge
                    else
                    {
                        // always Blue!
                        return Color.Blue;
                    }
                case "NiCd ULM":
                    // Discharge
                    if (test_type == "As Received" || test_type == "Capacity-1" || test_type == "Discharge" || test_type == "Custom Cap" || test_type == "")
                    {
                        Min4 = 1 * Cells;
                        Max = 1.05 * Cells;

                        if (Value > Max) { return Color.Green; }
                        else if (Value > Min4) { return Color.Orange; }
                        else { return Color.Red; }

                    }
                    // Charge
                    else
                    {
                        Min1 = 0.25 * Cells;
                        Min2 = 1.5 * Cells;
                        Min3 = 1.55 * Cells;
                        Max = ((-1 == (float)pci.Rows[station][7]) ? 1.82 : (float)pci.Rows[station][7]) * Cells;

                        if (Value > Max) { return Color.Red; }
                        else if (Value > Min3) { return Color.Green; }
                        else if (Value > Min2) { return Color.Orange; }
                        else if (Value > Min1) { return Color.Yellow; }
                        else { return Color.Red; }
                    }
            }

            // we'll return a purple if everything goes wrong
            return System.Drawing.Color.Purple;
        }

        //store the current test to compare to the old test
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
                    try
                    {
                        lockUpdate = true;
                        // first things first
                        // if the cscan isn't in use then lets just set the drop downs to the default and return...
                        if ((bool)d.Rows[currentRow][4] == false)
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                // just set to the cells readings..
                                comboBox2.Items.Clear();
                                toolStripComboBox2.ComboBox.Items.Clear();
                                comboBox3.Items.Clear();
                                toolStripComboBox3.ComboBox.Items.Clear();
                                radioButton1.Enabled = false;
                                radioButton2.Enabled = false;
                                updateR2(true);
                                comboBox2.Enabled = false;
                                toolStripComboBox2.ComboBox.Enabled = false;
                                comboBox3.Enabled = false;
                                toolStripComboBox3.ComboBox.Enabled = false;
                                radioButton2.Text = "Cells";
                                comboBox3.Items.Add("Cell Voltages");
                                toolStripComboBox3.ComboBox.Items.Add("Cell Voltages");
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
                            string tempWOS = d.Rows[currentRow][1].ToString();
                            char[] delims = { ' ' };
                            string[] A = tempWOS.Split(delims);
                            workOrder = A[0];
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
                                    toolStripComboBox2.ComboBox.Items.Clear();
                                    comboBox3.Items.Clear();
                                    toolStripComboBox3.ComboBox.Items.Clear();
                                    radioButton1.Enabled = false;
                                    radioButton2.Enabled = false;
                                    updateR2(true);
                                    comboBox2.Enabled = false;
                                    toolStripComboBox2.ComboBox.Enabled = false;
                                    comboBox3.Enabled = false;
                                    toolStripComboBox3.ComboBox.Enabled = false;
                                    if (GlobalVars.CScanData[currentRow] == null)
                                    {
                                        radioButton2.Text = "Cells";
                                        comboBox3.Items.Add("Cell Voltages");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell Voltages");
                                        updateC3("Cell Voltages");
                                    }
                                    else if (GlobalVars.CScanData[currentRow].CCID == 10)
                                    {
                                        radioButton2.Text = " ";
                                        comboBox3.Items.Add("Current Voltages");
                                        toolStripComboBox3.ComboBox.Items.Add("Current Voltages");
                                        updateC3("Current Voltages");
                                    }
                                    else
                                    {
                                        radioButton2.Text = "Cells";
                                        comboBox3.Items.Add("Cell Voltages");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell Voltages");
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
                                toolStripComboBox2.ComboBox.Enabled = false;
                                comboBox3.Enabled = false;
                                toolStripComboBox3.ComboBox.Enabled = false;
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
                                this.Invoke((MethodInvoker)delegate()
                                {
                                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                });
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
                                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                lockUpdate = false;
                                return;
                            }



                            string cellCable = GlobalVars.CScanData[currentRow].CCID.ToString();

                            //improve pick based on PCI here...
                            try
                            {
                                if ((int)pci.Rows[currentRow][3] != -1 && pci.Rows[currentRow][1].ToString() != "Sealed Lead Acid")
                                {
                                    if ((int)pci.Rows[currentRow][3] == 20)
                                    {
                                        cellCable = "1";
                                    }
                                    else if ((int)pci.Rows[currentRow][3] == 21)
                                    {
                                        cellCable = "21";
                                    }
                                    else if ((int)pci.Rows[currentRow][3] == 22)
                                    {
                                        cellCable = "3";
                                    }
                                }
                            }
                            catch
                            {
                                //didn't work...
                            }

                            this.Invoke((MethodInvoker)delegate()
                            {
                                switch (cellCable)
                                {
                                    case "1":
                                        // Battery combobox
                                        comboBox2.Items.Clear();
                                        toolStripComboBox2.ComboBox.Items.Clear();
                                        comboBox2.Items.Add("Voltage");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage");
                                        comboBox2.Items.Add("Current");
                                        toolStripComboBox2.ComboBox.Items.Add("Current");
                                        comboBox2.Items.Add("Temperature 1");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 1");
                                        comboBox2.Items.Add("Temperature 2");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 2");
                                        comboBox2.Items.Add("Temperature 3");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage");
                                        comboBox2.Items.Add("Temperature 4");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage");
                                        // Cells combobox
                                        comboBox3.Items.Clear();
                                        toolStripComboBox3.ComboBox.Items.Clear();
                                        comboBox3.Items.Add("Cell Voltages");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell Voltages");
                                        comboBox3.Items.Add("Cell 1");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 1");
                                        comboBox3.Items.Add("Cell 2");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 2");
                                        comboBox3.Items.Add("Cell 3");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 3");
                                        comboBox3.Items.Add("Cell 4");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 4");
                                        comboBox3.Items.Add("Cell 5");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 5");
                                        comboBox3.Items.Add("Cell 6");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 6");
                                        comboBox3.Items.Add("Cell 7");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 7");
                                        comboBox3.Items.Add("Cell 8");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 8");
                                        comboBox3.Items.Add("Cell 9");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 9");
                                        comboBox3.Items.Add("Cell 10");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 10");
                                        comboBox3.Items.Add("Cell 11");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 11");
                                        comboBox3.Items.Add("Cell 12");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 12");
                                        comboBox3.Items.Add("Cell 13");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 13");
                                        comboBox3.Items.Add("Cell 14");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 14");
                                        comboBox3.Items.Add("Cell 15");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 15");
                                        comboBox3.Items.Add("Cell 16");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 16");
                                        comboBox3.Items.Add("Cell 17");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 17");
                                        comboBox3.Items.Add("Cell 18");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 18");
                                        comboBox3.Items.Add("Cell 19");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 19");
                                        comboBox3.Items.Add("Cell 20");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 20");
                                        break;
                                    case "3":
                                        // Battery combobox
                                        comboBox2.Items.Clear();
                                        toolStripComboBox2.ComboBox.Items.Clear();
                                        comboBox2.Items.Add("Voltage 1");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage 1");
                                        comboBox2.Items.Add("Voltage 2");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage 2");
                                        comboBox2.Items.Add("Current");
                                        toolStripComboBox2.ComboBox.Items.Add("Current");
                                        comboBox2.Items.Add("Temperature 1");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 1");
                                        comboBox2.Items.Add("Temperature 2");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 2");
                                        comboBox2.Items.Add("Temperature 3");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 3");
                                        comboBox2.Items.Add("Temperature 4");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 4");
                                        // Cells combobox
                                        comboBox3.Items.Clear();
                                        toolStripComboBox3.ComboBox.Items.Clear();
                                        comboBox3.Items.Add("Cell Voltages");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell Voltages");
                                        comboBox3.Items.Add("Cell 1");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 1");
                                        comboBox3.Items.Add("Cell 2");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 2");
                                        comboBox3.Items.Add("Cell 3");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 3");
                                        comboBox3.Items.Add("Cell 4");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 4");
                                        comboBox3.Items.Add("Cell 5");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 5");
                                        comboBox3.Items.Add("Cell 6");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 6");
                                        comboBox3.Items.Add("Cell 7");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 7");
                                        comboBox3.Items.Add("Cell 8");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 8");
                                        comboBox3.Items.Add("Cell 9");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 9");
                                        comboBox3.Items.Add("Cell 10");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 10");
                                        comboBox3.Items.Add("Cell 11");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 11");
                                        comboBox3.Items.Add("Cell 12");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 12");
                                        comboBox3.Items.Add("Cell 13");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 13");
                                        comboBox3.Items.Add("Cell 14");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 14");
                                        comboBox3.Items.Add("Cell 15");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 15");
                                        comboBox3.Items.Add("Cell 16");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 16");
                                        comboBox3.Items.Add("Cell 17");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 17");
                                        comboBox3.Items.Add("Cell 18");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 18");
                                        comboBox3.Items.Add("Cell 19");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 19");
                                        comboBox3.Items.Add("Cell 20");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 20");
                                        comboBox3.Items.Add("Cell 21");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 21");
                                        comboBox3.Items.Add("Cell 22");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 22");
                                        break;
                                    case "4":
                                        // Battery combobox
                                        comboBox2.Items.Clear();
                                        toolStripComboBox2.ComboBox.Items.Clear();
                                        comboBox2.Items.Add("Voltage 1");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage 1");
                                        comboBox2.Items.Add("Voltage 2");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage 2");
                                        comboBox2.Items.Add("Voltage 3");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage 3");
                                        comboBox2.Items.Add("Current");
                                        toolStripComboBox2.ComboBox.Items.Add("Current");
                                        comboBox2.Items.Add("Temperature 1");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 1");
                                        comboBox2.Items.Add("Temperature 2");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 2");
                                        comboBox2.Items.Add("Temperature 3");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 3");
                                        comboBox2.Items.Add("Temperature 4");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 4");
                                        // Cells combobox
                                        comboBox3.Items.Clear();
                                        toolStripComboBox3.ComboBox.Items.Clear();
                                        comboBox3.Items.Add("Cell Voltages");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell Voltages");
                                        comboBox3.Items.Add("Cell 1");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 1");
                                        comboBox3.Items.Add("Cell 2");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 2");
                                        comboBox3.Items.Add("Cell 3");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 3");
                                        comboBox3.Items.Add("Cell 4");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 4");
                                        comboBox3.Items.Add("Cell 5");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 5");
                                        comboBox3.Items.Add("Cell 6");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 6");
                                        comboBox3.Items.Add("Cell 7");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 7");
                                        comboBox3.Items.Add("Cell 8");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 8");
                                        comboBox3.Items.Add("Cell 9");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 9");
                                        comboBox3.Items.Add("Cell 10");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 10");
                                        comboBox3.Items.Add("Cell 11");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 11");
                                        comboBox3.Items.Add("Cell 12");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 12");
                                        comboBox3.Items.Add("Cell 13");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 13");
                                        comboBox3.Items.Add("Cell 14");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 14");
                                        comboBox3.Items.Add("Cell 15");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 15");
                                        comboBox3.Items.Add("Cell 16");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 16");
                                        comboBox3.Items.Add("Cell 17");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 17");
                                        comboBox3.Items.Add("Cell 18");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 18");
                                        comboBox3.Items.Add("Cell 19");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 19");
                                        comboBox3.Items.Add("Cell 20");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 20");
                                        comboBox3.Items.Add("Cell 21");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 21");
                                        break;
                                    case "10":
                                        // Battery combobox
                                        comboBox2.Items.Clear();
                                        toolStripComboBox2.ComboBox.Items.Clear();
                                        comboBox2.Items.Add("Voltage 1");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage 1");
                                        comboBox2.Items.Add("Voltage 2");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage 2");
                                        comboBox2.Items.Add("Voltage 3");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage 3");
                                        comboBox2.Items.Add("Voltage 4");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage 4");
                                        comboBox2.Items.Add("Current");
                                        toolStripComboBox2.ComboBox.Items.Add("Current");
                                        comboBox2.Items.Add("Temperature 1");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 1");
                                        comboBox2.Items.Add("Temperature 2");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 2");
                                        comboBox2.Items.Add("Temperature 3");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 3");
                                        comboBox2.Items.Add("Temperature 4");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 4");
                                        // Cells combobox
                                        comboBox3.Items.Clear();
                                        toolStripComboBox3.ComboBox.Items.Clear();
                                        comboBox3.Items.Add("Current Voltages");
                                        toolStripComboBox3.ComboBox.Items.Add("Current Voltages");
                                        break;
                                    case "21":
                                        // Battery combobox
                                        comboBox2.Items.Clear();
                                        toolStripComboBox2.ComboBox.Items.Clear();
                                        comboBox2.Items.Add("Voltage");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage");
                                        comboBox2.Items.Add("Current");
                                        toolStripComboBox2.ComboBox.Items.Add("Current");
                                        comboBox2.Items.Add("Temperature 1");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 1");
                                        comboBox2.Items.Add("Temperature 2");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 2");
                                        comboBox2.Items.Add("Temperature 3");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 3");
                                        comboBox2.Items.Add("Temperature 4");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 4");
                                        // Cells combobox
                                        comboBox3.Items.Clear();
                                        toolStripComboBox3.ComboBox.Items.Clear();
                                        comboBox3.Items.Add("Cell Voltages");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell Voltages");
                                        comboBox3.Items.Add("Cell 1");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 1");
                                        comboBox3.Items.Add("Cell 2");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 2");
                                        comboBox3.Items.Add("Cell 3");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 3");
                                        comboBox3.Items.Add("Cell 4");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 4");
                                        comboBox3.Items.Add("Cell 5");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 5");
                                        comboBox3.Items.Add("Cell 6");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 6");
                                        comboBox3.Items.Add("Cell 7");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 7");
                                        comboBox3.Items.Add("Cell 8");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 8");
                                        comboBox3.Items.Add("Cell 9");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 9");
                                        comboBox3.Items.Add("Cell 10");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 10");
                                        comboBox3.Items.Add("Cell 11");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 11");
                                        comboBox3.Items.Add("Cell 12");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 12");
                                        comboBox3.Items.Add("Cell 13");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 13");
                                        comboBox3.Items.Add("Cell 14");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 14");
                                        comboBox3.Items.Add("Cell 15");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 15");
                                        comboBox3.Items.Add("Cell 16");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 16");
                                        comboBox3.Items.Add("Cell 17");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 17");
                                        comboBox3.Items.Add("Cell 18");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 18");
                                        comboBox3.Items.Add("Cell 19");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 19");
                                        comboBox3.Items.Add("Cell 20");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 20");
                                        comboBox3.Items.Add("Cell 21");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 21");
                                        break;
                                    default:
                                        // Battery combobox
                                        comboBox2.Items.Clear();
                                        toolStripComboBox2.ComboBox.Items.Clear();
                                        comboBox2.Items.Add("Voltage 1");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage 1");
                                        comboBox2.Items.Add("Voltage 2");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage 2");
                                        comboBox2.Items.Add("Voltage 3");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage 3");
                                        comboBox2.Items.Add("Voltage 4");
                                        toolStripComboBox2.ComboBox.Items.Add("Voltage 4");
                                        comboBox2.Items.Add("Current");
                                        toolStripComboBox2.ComboBox.Items.Add("Current");
                                        comboBox2.Items.Add("Temperature 1");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 1");
                                        comboBox2.Items.Add("Temperature 2");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 2");
                                        comboBox2.Items.Add("Temperature 3");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 3");
                                        comboBox2.Items.Add("Temperature 4");
                                        toolStripComboBox2.ComboBox.Items.Add("Temperature 4");
                                        // Cells combobox
                                        comboBox3.Items.Clear();
                                        toolStripComboBox3.ComboBox.Items.Clear();
                                        comboBox3.Items.Add("Cell Voltages");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell Voltages");
                                        comboBox3.Items.Add("Cell 1");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 1");
                                        comboBox3.Items.Add("Cell 2");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 2");
                                        comboBox3.Items.Add("Cell 3");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 3");
                                        comboBox3.Items.Add("Cell 4");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 4");
                                        comboBox3.Items.Add("Cell 5");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 5");
                                        comboBox3.Items.Add("Cell 6");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 6");
                                        comboBox3.Items.Add("Cell 7");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 7");
                                        comboBox3.Items.Add("Cell 8");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 8");
                                        comboBox3.Items.Add("Cell 9");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 9");
                                        comboBox3.Items.Add("Cell 10");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 10");
                                        comboBox3.Items.Add("Cell 11");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 11");
                                        comboBox3.Items.Add("Cell 12");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 12");
                                        comboBox3.Items.Add("Cell 13");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 13");
                                        comboBox3.Items.Add("Cell 14");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 14");
                                        comboBox3.Items.Add("Cell 15");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 15");
                                        comboBox3.Items.Add("Cell 16");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 16");
                                        comboBox3.Items.Add("Cell 17");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 17");
                                        comboBox3.Items.Add("Cell 18");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 18");
                                        comboBox3.Items.Add("Cell 19");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 19");
                                        comboBox3.Items.Add("Cell 20");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 20");
                                        comboBox3.Items.Add("Cell 21");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 21");
                                        comboBox3.Items.Add("Cell 22");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 22");
                                        comboBox3.Items.Add("Cell 23");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 23");
                                        comboBox3.Items.Add("Cell 24");
                                        toolStripComboBox3.ComboBox.Items.Add("Cell 24");
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
                                    toolStripComboBox3.ComboBox.SelectedIndex = 0;
                                    if (comboBox2.Text == "") 
                                    { 
                                        comboBox2.SelectedIndex = 0;
                                        toolStripComboBox2.ComboBox.SelectedIndex = 0;
                                    }
                                }
                                else
                                {
                                    comboBox2.SelectedIndex = 0;
                                    toolStripComboBox2.ComboBox.SelectedIndex = 0;
                                    updateC3(gs.Rows[currentRow][1].ToString());
                                    if (comboBox3.Text == "") 
                                    { 
                                        comboBox3.SelectedIndex = 0;
                                        toolStripComboBox3.ComboBox.SelectedIndex = 0;
                                    }
                                }
                                comboBox2.Enabled = true;
                                toolStripComboBox2.ComboBox.Enabled = true;
                                comboBox3.Enabled = true;
                                toolStripComboBox3.ComboBox.Enabled = true;

                                if (cellCable == "10") { radioButton2.Text = " "; }
                                else { radioButton2.Text = "Cells"; }
                                lockUpdate = false;

                            });// end invoke
                        }// end else
                    }// end try
                    catch
                    {
                        // throw it out...
                    }
                });// end helper thread
            

        }// end function

        // this will update the gs datatable when the radio buttons are changed
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            ThreadPool.QueueUserWorkItem(s =>
               {
                   this.Invoke((MethodInvoker)delegate()
                   {
                       if (comboBox3.Text != "") 
                       { 
                           gs.Rows[dataGridView1.CurrentRow.Index][0] = radioButton1.Checked; 
                       }
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
                               if (comboBox2.Text != "") 
                               { 
                                   gs.Rows[dataGridView1.CurrentRow.Index][1] = comboBox2.Text; 
                               }
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
                               if (comboBox3.Text != "") 
                               { 
                                   gs.Rows[dataGridView1.CurrentRow.Index][1] = comboBox3.Text; 
                               }
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
                toolStripComboBox2.ComboBox.Text = inVal;
            }
        }

        private readonly object combo3Lock = new object();

        private void updateC3(string inVal)
        {
            lock (combo3Lock)
            {
                comboBox3.Text = inVal;
                toolStripComboBox3.ComboBox.Text = inVal;
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

        private static double ConvertCelsiusToFahrenheit(double c)
        {
            return ((9.0 / 5.0) * c) + 32;
        }

    }
}
