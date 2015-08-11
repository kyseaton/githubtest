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
        public SerialPort ICComPort = new SerialPort();

        public CancellationTokenSource cPollIC;

        //for checking the IC when selected
        bool check =  false;
        int toCheck;
        int chanNum;

        //for critical operations (Start,Stop, etc)
        bool [] criticalNum = new bool[16];

        



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

                            if ((bool) d.Rows[j][8] && (bool) d.Rows[j][4] && (string) d.Rows[j][9] != "" && (string) d.Rows[j][10] == "ICA")
                            {
                                Thread.Sleep(500);
                                try
                                {
                                    // send the short command based on the settings for the charger...
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
                                            updateD(j,11,testData.runStatus);    
                                        }
                                        else
                                        {
                                            updateD(j,11,"offline!"); 
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
                                    updateD(j,11,"");
                                } 
                            }

                            if (check)
                            {
                                Thread.Sleep(500);
                                try
                                {
                                    // send the short command based on the settings for the charger...
                                    ICComPort.Write(GlobalVars.ICSettings[toCheck].outText, 0, 28);
                                    // wait for a response
                                    tempBuff = ICComPort.ReadTo("Z");
                                    // if we got one then we can determine that we have an ICA
                                    updateD(chanNum,10,"ICA");
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
                                    else { throw ex; }
                                }       // end catch
                            }       // end else if

                            // we need to check for critical operation also!
                            for(int i = 0;i < 16; i++)
                            {
                                if (criticalNum[i] == true)
                                {
                                    try
                                    {
                                        Thread.Sleep(500);
                                        // send the short command based on the settings for the charger...
                                        ICComPort.Write(GlobalVars.ICSettings[i].outText, 0, 28);
                                        // wait for a response
                                        tempBuff = ICComPort.ReadTo("Z");
                                        // and we don't need to check any more
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            rtbIncoming.Text = "Critical  " + i.ToString() + "  :  " + tempBuff;
                                        });
                                        Thread.Sleep(200);
                                    }
                                    catch (Exception ex)
                                    {
                                        if (ex is System.TimeoutException)
                                        {
                                            Thread.Sleep(100);
                                        }
                                        else { throw ex; }
                                    }       // end catch
                                } // end if
                            }// end for

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
                Thread.Sleep(3000);
                // now we'll make sure we're not looking anymore...
                check = false;
            }); // end thread


        }


    }
}
