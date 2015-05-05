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

            string tempBuff;
            CScanDataStore testData = new CScanDataStore();

            // Open the comport
            ICComPort.ReadTimeout = 500;
            ICComPort.PortName = Convert.ToString(comboBox9.Text);
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
                        rtbIncoming.Text = "";
                        rtbIncoming.Text = tempBuff;
                        button6.Enabled = true;
                    });

                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    ICComPort.Close();
                    this.Invoke((MethodInvoker)delegate
                    {
                        button5.Enabled = true;
                    });
                    
                }
            });

        }


        private void SetText(string text)
        {
            this.rtbIncoming.Text += text;
        }


    }
}
