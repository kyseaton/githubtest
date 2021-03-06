﻿using System;
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
        /// 

        

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

            // First do a commPort Check
            try
            {
                CSCANComPort = new SerialPort();
                ICComPort = new SerialPort();
                CSCANComPort.PortName = GlobalVars.CSCANComPort;
                ICComPort.PortName = GlobalVars.ICComPort;
                
                CSCANComPort.Open();
                ICComPort.Open();
            }
            catch
            {
                label8.Text = "Check Comports' Settings";
                label8.Visible = true;
                //MessageBox.Show("There is a Comport configuration issue.  Please check the Data Hub connections and the ComPort Settings");
                return;
            }
            finally
            {
                CSCANComPort.Close();
                CSCANComPort.Dispose();
                ICComPort.Close();
                ICComPort.Dispose();

            }



            // to make this all more readable...
            pollCScans();
            pollICs();
            sequentialScan();


        }                           // end Scan()

    }
}
