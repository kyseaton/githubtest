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

        CancellationTokenSource cPollIC = new CancellationTokenSource();



        public void pollICs()
        {
            //This code is here to poll the ICs/////////////////////////////////////////////////////////////////////////////////////
            // We are going to run all of this code on a helper thread, as to improve GUI performance//////////////////////////////////////////

            string tempBuff = "";
            ICDataStore testData;

            // Open the comport
            ICComPort.ReadTimeout = 100;
            ICComPort.PortName = "COM11";
            ICComPort.BaudRate = 19200;
            ICComPort.DataBits = 8;
            ICComPort.StopBits = StopBits.One;
            ICComPort.Handshake = Handshake.None;
            ICComPort.Parity = Parity.None;
            ICComPort.Open();

            //OUTPUT DATA STRUCTURE ----------------------------------
            byte T1 = 0;
            byte T2 = 1;
            //--- P ---
            byte KE1 = 0;                                           //KE1 is defined at the beginning of this subroutine: 0=query, 1=data, 2=command, 3=data/command
            byte KE2 = 0;                                           //KE2 not used for now (3 bits, available)
            byte KE3 = 0;                                           //action: 0=clear, 1=run, 2=stop, 3=reset
            byte KM0 = (byte) ((KE1 + 4 * KE2 + 64 * KE3));    //first command [type + test + action] lower 8 bits
            byte KM1 = (byte) (10);                            //'Mode (10, 11, 12, 21, 21, 30, 31, 32 [for now] )

            //--- A ---
            byte KM2 = (byte) (0);                             //CT1H, Charge Time 1, Hours
            byte KM3 = (byte) (0);                             //CT1M, Charge Time 1, Minutes
            byte KM4 = (byte)(0);                             //CC1H, Charge current 1, High (byte)
            byte KM5 = (byte)(0);                             //CC1L, Charge current 1, Low (byte)
            byte KM6 = (byte)(0);                             //CV1H, Charge voltage 1, High (byte)
            byte KM7 = (byte)(0);                             //CV1L, Charge voltage 1, Low (byte)

            //--- B ---
            byte KM8 = (byte)(0);                             //CT2H, Charge Time 2, Hours
            byte KM9 = (byte)(0);                             //CT2M, Charge Time 2, Minutes
            byte KM10 = (byte)(0);                             //CC2H, Charge Current 2, High (byte)
            byte KM11 = (byte)(0);                            //CC2L, Charge Current 2, Low (byte)
            byte KM12 = (byte)(0);                            //CV2H, Charge Voltage 2, High (byte)
            byte KM13 = (byte)(0);                            //CV2L, Charge voltage 2, Low (byte)

            //--- C ---
            byte KM14 = (byte)(0);                            //DTH, Discharge Time, Hours
            byte KM15 = (byte)(0);                            //DTM, Discharge Time, Minutes
            byte KM16 = (byte)(0);                            //DCH, Discharge current, High (byte)
            byte KM17 = (byte)(0);                            //DCL, Discharge current, Low (byte)
            byte KM18 = (byte)(0);                            //DVH, Discharge voltage, High (byte)
            byte KM19 = (byte)(0);                            //DVL, Discharge Voltage, Low (byte)
            byte KM20 = (byte)(0);                            //CRH, Discharge Resistance, High (byte)
            byte KM21 = (byte)(0);                            //CRL, Discharge Resistance, Low (byte)

            byte ULCH = 0;                                          //---

            string outText = "";



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
                            Thread.Sleep(100);
                            //putting the cancellation token in a often looked at place...
                            if (token.IsCancellationRequested) return;

                            if ((bool) d.Rows[j][8] && (bool) d.Rows[j][4])
                            {
                                Thread.Sleep(900);
                                try
                                {
                                    // do this all on a threadpool thread

                                    // send the polling command
                                    outText = "~" + j.ToString("00") + "L0000000000000000000000Z";    //send wake-up character, terminal ID, WDO, commands and Z

                                    ICComPort.Write(outText);
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
                                        d.Rows[j][11] = testData.runStatus;
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
                                            rtbIncoming.Text = "No Dice!!! " + outText + System.Environment.NewLine + tempBuff;
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


        private void SetText(string text)
        {
            this.rtbIncoming.Text += text;
        }


    }
}
