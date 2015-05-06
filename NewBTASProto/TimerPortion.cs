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
using System.Timers;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;

namespace NewBTASProto
{
    
    public delegate void timerTick();

    public partial class Main_Form : Form
    {
        // This is the code to update the time label///////////////////////////////////////////////////////////////////////////////////////

        /// <summary>
        /// This is a timer running on a helper thread (Prevents lag due to UI)
        /// </summary>
        private System.Timers.Timer tmrTimersTimer;
        /// <summary>
        /// This is the delegate for timer events
        /// </summary>

        private void InitializeTimers()
        {

            //Initialize System.Timers.Timer (this type is safe in multi-threaded apps)...
            tmrTimersTimer = new System.Timers.Timer();
            tmrTimersTimer.Interval = 500;
            tmrTimersTimer.Elapsed += new ElapsedEventHandler(tmrTimersTimer_Elapsed);
            tmrTimersTimer.Start(); 

        }

        private void tmrTimersTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                //Call a delegate instantance on the UI thread using Invoke
                timeLabel.Invoke(new timerTick(this.UpdateGrid));
            }
            catch
            {
                //do nothing
            }

        }

        public void UpdateGrid(){
            //Update the time label
            timeLabel.Text = System.DateTime.Now.ToString("HH:mm:ss");
            dateLabel.Text = System.DateTime.Now.ToString("MM/dd/yyyy");
        }

        //This code is here to poll the CSCANs/////////////////////////////////////////////////////////////////////////////////////
        // We are going to run all of this code on a helper thread, as to improve GUI performance//////////////////////////////////////////
        public void Scan()
        {

            string tempBuff;
            CScanDataStore testData;

            //MOVE TO A STARTUP LOCATION!!!!!!!!!!!!!!!!!!!!!
            // Open the comport
            ComPort.ReadTimeout = 200;
            ComPort.PortName = Convert.ToString(cboPorts.Text);
            ComPort.BaudRate = 19200;
            ComPort.DataBits = 8;
            ComPort.StopBits = StopBits.One;
            ComPort.Handshake = Handshake.None;
            ComPort.Parity = Parity.None;

            ThreadPool.QueueUserWorkItem(_ =>
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
                                    ComPort.Open();
                                    // do this all on a threadpool thread

                                    // send the polling command
                                    string outText = ("~" + (dataGridView1.CurrentRow.Index + 16).ToString("00") + tbWDO.Text.ToString() + tbKM1.Text.ToString() + tbKM2.Text.ToString() + "Z");
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
                                            textBox1.Text += "Cell #" + (i + 1).ToString() + ":  " + testData.orderedCells[i].ToString("0.000") + Environment.NewLine;
                                        }
                                        textBox1.Text += "Temp Plate 1:  " + testData.TP1 + Environment.NewLine;
                                        textBox1.Text += "Temp Plate 2:  " + testData.TP2 + Environment.NewLine;
                                        textBox1.Text += "Temp Plate 3:  " + testData.TP3 + Environment.NewLine;
                                        textBox1.Text += "Temp Plate 4:  " + testData.TP4 + Environment.NewLine;
                                        textBox1.Text += "Ambient Temp:  " + testData.TP5 + Environment.NewLine;
                                        textBox1.Text += "Reference:  " + testData.ref95V + Environment.NewLine;

                                    });
                                }
                                catch (Exception ex)
                                {
                                    ComPort.Close();
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
                                    else
                                    {
                                        throw ex;
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
                                    ComPort.Open();
                                    // do this all on a threadpool thread

                                    // send the polling command
                                    string outText = ("~" + (j + 16).ToString("00") + tbWDO.Text.ToString() + tbKM1.Text.ToString() + tbKM2.Text.ToString() + "Z");
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
                                        // first set the cell to green
                                        dataGridView1.Rows[j].Cells[4].Style.BackColor = Color.Green;
                                    });


                                }
                                catch (Exception ex)
                                {
                                    ComPort.Close();
                                    if (ex is System.TimeoutException)
                                    {
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            dataGridView1.Rows[j].Cells[4].Style.BackColor = Color.Red;
                                        });
                                    }
                                    else
                                    {
                                        throw ex;
                                    }
                                }
                            }           // end if
                            else if (j != dataGridView1.CurrentRow.Index || (bool)d.Rows[j][4] == false)
                            {
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        dataGridView1.Rows[j].Cells[4].Style.BackColor = Color.Gainsboro;
                                    });

                            }
                        }               // end for
                    }                   // end while
                }                       // end try
                catch(Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }                   // end catch
            });                     // end thread


            // to make this all more readable...

            pollIC();

        //This code is here to poll the ICs/////////////////////////////////////////////////////////////////////////////////////
        // We are going to run all of this code on a helper thread, as to improve GUI performance/////////////////////////////////////////


        }                           // end Scan()

    }
}
