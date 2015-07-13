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

namespace NewBTASProto
{
    public partial class Main_Form : Form
    {

        

        //ComPort.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(port_DataReceived_1);
        //EventArgs E = new EventArgs();

        /// <summary>
        /// Serial Stuff defined here
        /// </summary>
        SerialPort CSCANComPort = new SerialPort();

        internal delegate void SerialDataReceivedEventHandlerDelegate(object sender, SerialDataReceivedEventArgs e);
        delegate void SetTextCallback(string text);
        string InputData = String.Empty;

        // this is for the chart on the main form
        DataSet graphMainSet = new DataSet();

        // cancellation token
        CancellationTokenSource cPollCScans = new CancellationTokenSource();
        CancellationTokenSource sequentialScanT = new CancellationTokenSource();


        public void pollCScans()
        {

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

                            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][4])
                            {

                                try
                                {
                                    CSCANComPort.Open();
                                    // do this all on a threadpool thread

                                    // send the polling command
                                    string outText = "~" + (dataGridView1.CurrentRow.Index + 16).ToString("00") + "L00Z";
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

                                    if ((bool)d.Rows[dataGridView1.CurrentRow.Index][4] && tempClick == dataGridView1.CurrentRow.Index)  // test to see if we've clicked in the mean time...
                                    {
                                        //put this new data in the chart!
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            // first set the cell to green
                                            dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[4].Style.BackColor = Color.Green;

                                            //if that row is selected, update the chart portion
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

                                            for (int i = 0; i < 24; i++)
                                            {
                                                series1.Points.AddXY(i + 1, testData.orderedCells[i]);
                                                // color test
                                                series1.Points[i].Color = pointColorMain(0, 1, testData.orderedCells[23 - i], 4);
                                            }
                                            chart1.Invalidate();
                                            chart1.ChartAreas[0].RecalculateAxesScale();

                                            //Real Time Data Portion
                                            string tempText = "";
                                            tempText = System.DateTime.Now.ToString("M/d/yyyy") + "       Terminal:  " + testData.terminalID.ToString() + Environment.NewLine;
                                            tempText += "Temp. Cable:  " + testData.TCAB.ToString() + "   (" + testData.tempPlateType + ")" + Environment.NewLine;
                                            tempText += "Cells Cable:  " + testData.CCID.ToString() + "   (" + testData.cellCableType + ")" + Environment.NewLine;
                                            tempText += "Shunt Cable:  " + testData.SHCID.ToString() + "   (" + testData.shuntCableType + ")" + Environment.NewLine;
                                            tempText += "Voltage Batt 1:  " + testData.VB1.ToString("00.00") + Environment.NewLine;
                                            tempText += "Voltage Batt 2:  " + testData.VB2.ToString("00.00") + Environment.NewLine;
                                            tempText += "Voltage Batt 3:  " + testData.VB3.ToString("00.00") + Environment.NewLine;
                                            tempText += "Voltage Batt 4:  " + testData.VB4.ToString("00.00") + Environment.NewLine;
                                            tempText += "Current#1:  " + testData.currentOne.ToString("00.00") + Environment.NewLine;
                                            tempText += "Current#2:  " + testData.currentTwo.ToString("00.00") + Environment.NewLine;

                                            for (int i = 0; i < 24; i++)
                                            {
                                                tempText += "Cell #" + (i + 1).ToString() + ":  " + testData.orderedCells[i].ToString("0.000") + Environment.NewLine;
                                            }
                                            tempText += "Temp Plate 1:  " + testData.TP1 + Environment.NewLine;
                                            tempText += "Temp Plate 2:  " + testData.TP2 + Environment.NewLine;
                                            tempText += "Temp Plate 3:  " + testData.TP3 + Environment.NewLine;
                                            tempText += "Temp Plate 4:  " + testData.TP4 + Environment.NewLine;
                                            tempText += "Ambient Temp:  " + testData.TP5 + Environment.NewLine;
                                            tempText += "Reference:  " + testData.ref95V + Environment.NewLine;
                                            tempText += "Program Version " + testData.programVersion;
                                            textBox1.Text = tempText;


                                        });
                                    }
                                }
                                catch (Exception ex)
                                {
                                    if (ex is System.TimeoutException)
                                    {
                                        // make sure there haven't been any clicks in the mean time...
                                        if ((bool)d.Rows[j][4] && tempClick == dataGridView1.CurrentRow.Index)
                                        {
                                            this.Invoke((MethodInvoker)delegate
                                            {

                                                dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[4].Style.BackColor = Color.Red;
                                                chart1.Series.Clear();
                                                chart1.Invalidate();
                                                textBox1.Text = "";
                                            });
                                        }
                                        CSCANComPort.Close();

                                    }
                                }
                            }// end if for selected case
                            


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
                                        // do this all on a threadpool thread

                                        // send the polling command
                                        string outText = ("~" + (j + 16).ToString("00") + "L00Z");
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

                                        if ((bool)d.Rows[j][4])  // added to help with gui look
                                        {
                                            //put this new data in the chart!
                                            this.Invoke((MethodInvoker)delegate
                                            {
                                                // set the cell to green
                                                dataGridView1.Rows[j].Cells[4].Style.BackColor = Color.Green;
                                            });
                                        }  // end if
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
                                                    dataGridView1.Rows[j].Cells[4].Style.BackColor = Color.Red;
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

                            }
                        }               // end for
                        // end while
                    }                       // end try
                    catch (Exception ex)
                    {
                        if (token.IsCancellationRequested) return;
                        else
                        {
                            MessageBox.Show(ex.ToString());
                        }
                        
                    }                   // end catch
                }                       // end while
            }, cPollCScans.Token);                     // end thread

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
                case 4:
                    Min1 = 0.1;
                    Min2 = 1.2;
                    Min3 = 1.25;
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



    }
}
