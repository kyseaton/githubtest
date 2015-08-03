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
using System.Text;

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

                            if ((bool)d.Rows[j][8] && (bool)d.Rows[j][4] && d.Rows[j][9] != "" )
                            {
                                Thread.Sleep(500);
                                try
                                {
                                    // send command based on the settings for the charger...
                                    ICComPort.Write(GlobalVars.ICSettings[Convert.ToInt32(d.Rows[j][9])].outText, 0, 28);
                                    // wait for a response
                                    tempBuff = ICComPort.ReadTo("Z");
                                    //do something with the new data
                                    char[] delims = { ' ' };
                                    string[] A = tempBuff.Split(delims);
                                    //A[1] has the terminal ID in it
                                    testData = new ICDataStore(A);

                                    //put this new data in the chart...
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        if (testData.online == true)
                                        {
                                            d.Rows[j][11] = testData.runStatus;    
                                        }
                                        else
                                        {
                                            d.Rows[j][11] = "offline!"; 
                                        }
                                        
                                        rtbIncoming.Text = j.ToString() + "  :  " + tempBuff;
                                        
                                    });

                                    Thread.Sleep(200);


                                }
                                catch (Exception ex)
                                {
                                    if (ex is System.TimeoutException)
                                    {
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            tempBuff = ICComPort.ReadExisting();
                                            rtbIncoming.Text = "Com Error" + System.Environment.NewLine + tempBuff;
                                        });
                                        Thread.Sleep(100);
                                    }
                                    else { throw ex; }
                                }       // end catch
                            }       // end if
                            else 
                            {
                                if ((string) d.Rows[j][11] != "")
                                {
                                    d.Rows[j][11] = "";
                                } 
                            }
                        }           // end for
                    }               // end try
                    catch (Exception ex)
                    {
                        if (token.IsCancellationRequested) return;
                        else
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }

                }                   // end while
            },cPollIC.Token); // end thread


        }  // end pollICs


    }
}
