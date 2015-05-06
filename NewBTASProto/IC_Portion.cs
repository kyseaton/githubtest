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
        SerialPort ICComPort = new SerialPort();

        internal delegate void SerialDataReceivedEventHandlerDelegate2(object sender, SerialDataReceivedEventArgs e);
        delegate void SetTextCallback2(string text);



        public void pollIC()
        {
            //This code is here to poll the ICs/////////////////////////////////////////////////////////////////////////////////////
            // We are going to run all of this code on a helper thread, as to improve GUI performance//////////////////////////////////////////

            string tempBuff;
            ICDataStore testData;

            // Open the comport
            ICComPort.ReadTimeout = 300;
            ICComPort.PortName = "COM11";
            ICComPort.BaudRate = 19200;
            ICComPort.DataBits = 8;
            ICComPort.StopBits = StopBits.One;
            ICComPort.Handshake = Handshake.None;
            ICComPort.Parity = Parity.None;

            //OUTPUT DATA STRUCTURE ----------------------------------
            byte T1 = 0;
            byte T2 = 1;
            //--- P ---
            byte KE1 = 0;                                           //KE1 is defined at the beginning of this subroutine: 0=query, 1=data, 2=command, 3=data/command
            byte KE2 = 0;                                           //KE2 not used for now (3 bits, available)
            byte KE3 = 0;                                           //action: 0=clear, 1=run, 2=stop, 3=reset
            byte KM0 = (byte) ((KE1 + 4 * KE2 + 64 * KE3) + 48);    //first command [type + test + action] lower 8 bits
            byte KM1 = (byte) (10 + 48);                            //'Mode (10, 11, 12, 21, 21, 30, 31, 32 [for now] )

            //--- A ---
            byte KM2 = (byte) (0 + 48);                             //CT1H, Charge Time 1, Hours
            byte KM3 = (byte) (0 + 48);                             //CT1M, Charge Time 1, Minutes
            byte KM4 = (byte) (0 + 48);                             //CC1H, Charge current 1, High (byte)
            byte KM5 = (byte) (0 + 48);                             //CC1L, Charge current 1, Low (byte)
            byte KM6 = (byte) (0 + 48);                             //CV1H, Charge voltage 1, High (byte)
            byte KM7 = (byte) (0 + 48);                             //CV1L, Charge voltage 1, Low (byte)

            //--- B ---
            byte KM8 = (byte) (0 + 48);                             //CT2H, Charge Time 2, Hours
            byte KM9 = (byte) (0 + 48);                             //CT2M, Charge Time 2, Minutes
            byte KM10 = (byte) (0 + 48);                             //CC2H, Charge Current 2, High (byte)
            byte KM11 = (byte) (0 + 48);                            //CC2L, Charge Current 2, Low (byte)
            byte KM12 = (byte) (0 + 48);                            //CV2H, Charge Voltage 2, High (byte)
            byte KM13 = (byte) (0 + 48);                            //CV2L, Charge voltage 2, Low (byte)

            //--- C ---
            byte KM14 = (byte) (0 + 48);                            //DTH, Discharge Time, Hours
            byte KM15 = (byte) (0 + 48);                            //DTM, Discharge Time, Minutes
            byte KM16 = (byte) (0 + 48);                            //DCH, Discharge current, High (byte)
            byte KM17 = (byte) (0 + 48);                            //DCL, Discharge current, Low (byte)
            byte KM18 = (byte) (0 + 48);                            //DVH, Discharge voltage, High (byte)
            byte KM19 = (byte) (0 + 48);                            //DVL, Discharge Voltage, Low (byte)
            byte KM20 = (byte) (0 + 48);                            //CRH, Discharge Resistance, High (byte)
            byte KM21 = (byte) (0 + 48);                            //CRL, Discharge Resistance, Low (byte)

            byte ULCH = 0;                                          //---



            ThreadPool.QueueUserWorkItem(_ =>
            {
                while (true)
                {
                    for (int j = 0; j < 16; j++)
                    {



                        try
                        {
                            ICComPort.Open();
                            // do this all on a threadpool thread

                            // send the polling command
                            string outText = "~" + T1.ToString() + T2.ToString() + "L" + KM0 + KM1 + KM2 + KM3 + KM4 + KM5 + KM6 + KM7 + KM8 + KM9 + KM10 + KM11 + KM12 + KM13 + KM14 + KM15 + KM16 + KM17 + KM18 + KM19 + KM20 + KM21 + ULCH + "Z";    //send wake-up character, terminal ID, WDO, commands and Z

                            ICComPort.Write(outText);
                            // wait for a response
                            tempBuff = ICComPort.ReadTo("Z");
                            // close the comport
                            ICComPort.Close();
                            //do something with the new data
                            char[] delims = { ' ' };
                            string[] A = tempBuff.Split(delims);
                            //A[1] has the terminal ID in it
                            ICDataStore currentICData = new ICDataStore(A);

                            //put this new data in the chart!
                            this.Invoke((MethodInvoker)delegate
                            {
                                rtbIncoming.Text = j.ToString() +"  :";
                                rtbIncoming.Text = tempBuff;
                                button6.Enabled = true;
                            });


                        }
                        catch (Exception ex)
                        {
                            ICComPort.Close();
                            if (ex is System.TimeoutException)
                            {
                                // serial port timed out...
                            }
                        }
                    }
                }
            }); // end thread

/*            ThreadPool.QueueUserWorkItem(_ =>
            {
                // this function is a while loop that is going to run forever
                try
                {

                    while (true)
                    {
                        for (int j = 0; j < 16; j++)
                        {
                            // first look at the selected row and then recheck all the other rows...
                            // look for the "In Use" columns and check for attached cscans
                            if ((bool)d.Rows[dataGridView1.CurrentRow.Index][4])
                            {
                                try
                                {
                                    ICComPort.Open();
                                    // do this all on a threadpool thread

                                    // send the polling command
                                    string outText = ("~" + (dataGridView1.CurrentRow.Index + 16).ToString("00") + tbWDO.Text.ToString() + tbKM1.Text.ToString() + tbKM2.Text.ToString() + "Z");
                                    ICComPort.Write(outText);
                                    // wait for a response
                                    tempBuff = ICComPort.ReadTo("Z");
                                    // close the comport
                                    ICComPort.Close();
                                    //do something with the new data
                                    char[] delims = { ' ' };
                                    string[] A = tempBuff.Split(delims);
                                    //A[1] has the terminal ID in it
                                    testData = new CScanDataStore(A);

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



                                    });
                                }
                                catch (Exception ex)
                                {
                                    ICComPort.Close();
                                    if (ex is System.TimeoutException)
                                    {
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[4].Style.BackColor = Color.Red;
                                            chart1.Series.Clear();
                                            chart1.Invalidate();
                                            textBox1.Text = "";
                                        });

                                    }
                                }
                            }// end if for selected case
                            else
                            {
                                this.Invoke((MethodInvoker)delegate
                                {
                                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[4].Style.BackColor = Color.Gainsboro;
                                    chart1.Series.Clear();
                                    chart1.Invalidate();
                                    textBox1.Text = "";
                                });

                            }

                            // now look at all of the other cases to up date the label

                            if ((bool)d.Rows[j][4] && j != dataGridView1.CurrentRow.Index)
                            {

                                // look for the "In Use" columns and check for attached cscans
                                try
                                {
                                    ICComPort.Open();
                                    // do this all on a threadpool thread

                                    // send the polling command
                                    string outText = ("~" + (j + 16).ToString("00") + tbWDO.Text.ToString() + tbKM1.Text.ToString() + tbKM2.Text.ToString() + "Z");
                                    ICComPort.Write(outText);
                                    // wait for a response
                                    tempBuff = ICComPort.ReadTo("Z");
                                    // close the comport
                                    ICComPort.Close();
                                    //do something with the new data
                                    char[] delims = { ' ' };
                                    string[] A = tempBuff.Split(delims);
                                    //A[1] has the terminal ID in it
                                    testData = new CScanDataStore(A);


                                    //put this new data in the chart!
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        // first set the cell to green
                                        dataGridView1.Rows[j].Cells[4].Style.BackColor = Color.Green;
                                    });


                                }
                                catch (Exception ex)
                                {
                                    ICComPort.Close();
                                    if (ex is System.TimeoutException)
                                    {
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            dataGridView1.Rows[j].Cells[4].Style.BackColor = Color.Red;
                                        });
                                    }
                                }
                            }           // end if
                            else if (j != dataGridView1.CurrentRow.Index || (bool)d.Rows[j][4] == false)
                            {
                                try
                                {
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        dataGridView1.Rows[j].Cells[4].Style.BackColor = Color.Gainsboro;
                                    });
                                }
                                catch
                                {

                                }

                            }
                        }               // end for
                    }                   // end while
                }                       // end try
                catch
                {

                }                   // end catch
            });                     // end thread
 * */

        }  // end pollICs


        private void SetText(string text)
        {
            this.rtbIncoming.Text += text;
        }


    }
}
