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

namespace NewBTASProto
{
    public partial class Main_Form : Form
    {

        //ComPort.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(port_DataReceived_1);
        //EventArgs E = new EventArgs();

        /// <summary>
        /// Serial Stuff defined here
        /// </summary>
        SerialPort ComPort = new SerialPort();

        internal delegate void SerialDataReceivedEventHandlerDelegate(object sender, SerialDataReceivedEventArgs e);
        delegate void SetTextCallback(string text);
        string InputData = String.Empty;

        // this is for the chart on the main form
        DataSet graphMainSet = new DataSet();


        public void pollCScan(int terminalID)
        {

            string tempBuff;
            CScanDataStore testData = new CScanDataStore();

            // Open the comport
            ComPort.ReadTimeout = 10000;
            ComPort.PortName = Convert.ToString(cboPorts.Text);
            ComPort.BaudRate = Convert.ToInt32(cboBaudRate.Text);
            ComPort.DataBits = Convert.ToInt32(cboDataBits.Text);
            ComPort.StopBits = (StopBits)Enum.Parse(typeof(StopBits), cboStopBits.Text);
            ComPort.Handshake = (Handshake)Enum.Parse(typeof(Handshake), cboHandShaking.Text);
            ComPort.Parity = (Parity)Enum.Parse(typeof(Parity), cboParity.Text);

            ThreadPool.QueueUserWorkItem(_ =>
            {
                

                try
                {
                    ComPort.Open();
                    // do this all on a threadpool thread

                    // send the polling command
                    string outText = ("~" + tbT1.Text.ToString() + tbT2.Text.ToString() + tbWDO.Text.ToString() + tbKM1.Text.ToString() + tbKM2.Text.ToString() + "Z");
                    ComPort.Write(outText);
                    // wait for a response
                    tempBuff = ComPort.ReadTo("Z");
                    // close the comport
                    ComPort.Close();
                    //do something with the new data
                    char[] delims = { ' ' };
                    string[] A = tempBuff.Split(delims);
                    //A[1] has the terminal ID in it
                    testData = new CScanDataStore(A);
                    
                    //put this new data in the chart!
                    this.Invoke((MethodInvoker)delegate
                    {
                        //chart portion
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
                            series1.Points.AddXY(i + 1, testData.orderedCells[23-i]);
                            // color test
                            series1.Points[i].Color = pointColorMain(0, 1,testData.orderedCells[23-i], 4);
                        }
                        chart1.Invalidate();
                        chart1.ChartAreas[0].RecalculateAxesScale();

                        //Real Time Data Portion
                        textBox1.Text = System.DateTime.Now.ToString("M/d/yyyy") + "       Terminal:  " + testData.terminalID.ToString() + Environment.NewLine;
                        textBox1.Text += "Temp. Cable:  " + testData.TCAB.ToString() + "   (" + testData.tempPlateType + ")" + Environment.NewLine;
                        textBox1.Text += "Cells Cable:  " + testData.CCID.ToString() + "   (" + testData.cellCableType + ")" + Environment.NewLine;
                        textBox1.Text += "Shunt Cable:  " + testData.SHCID.ToString() + "   (" + testData.shuntCableType + ")" + Environment.NewLine;
                        textBox1.Text += "Voltage Batt 1:  " + testData.VB1.ToString("00.00") + Environment.NewLine;
                        textBox1.Text += "Voltage Batt 2:  " + testData.VB2.ToString("00.00") + Environment.NewLine;
                        textBox1.Text += "Voltage Batt 3:  " + testData.VB3.ToString("00.00") + Environment.NewLine;
                        textBox1.Text += "Voltage Batt 4:  " + testData.VB4.ToString("00.00") + Environment.NewLine;
                        textBox1.Text += "Current#1:  " + testData.currentOne.ToString("00.00") + Environment.NewLine;
                        textBox1.Text += "Current#2:  " + testData.currentTwo.ToString("00.00") + Environment.NewLine;
                        for (int i = 0; i < 24; i++)
                        {
                            textBox1.Text += "Cell #" + i.ToString()+":  " + testData.orderedCells[23-i].ToString("0.000") + Environment.NewLine;
                        }
                        textBox1.Text += "Temp Plate 1:  " + testData.TP1 + Environment.NewLine;
                        textBox1.Text += "Temp Plate 2:  " + testData.TP2 + Environment.NewLine;
                        textBox1.Text += "Temp Plate 3:  " + testData.TP3 + Environment.NewLine;
                        textBox1.Text += "Temp Plate 4:  " + testData.TP4 + Environment.NewLine;
                        textBox1.Text += "Ambient Temp:  " + testData.TP5 + Environment.NewLine;
                        textBox1.Text += "Reference:  " + testData.ref95V + Environment.NewLine;

                            
                            

                        button5.Enabled = true;
                    });

                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    ComPort.Close();
                    this.Invoke((MethodInvoker)delegate
                    {
                        button5.Enabled = true;
                    });
                    
                }
            });

        }

        private void btnGetSerialPorts_Click_Click(object sender, EventArgs e)
        {
            
        }

        private void port_DataReceived_1(object sender, SerialDataReceivedEventArgs e)
        {
            InputData = ComPort.ReadExisting();
            if (InputData != String.Empty)
            {
                this.BeginInvoke(new SetTextCallback(SetText), new object[] { InputData });
            }
        }

        private void SetText(string text)
        {
            this.rtbIncoming.Text += text;
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
