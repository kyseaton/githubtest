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

namespace NewBTASProto
{
    public delegate void timerTick();

    public partial class Main_Form : Form
    {

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
            tmrTimersTimer.Interval = 1000;
            tmrTimersTimer.Elapsed += new ElapsedEventHandler(tmrTimersTimer_Elapsed);
            tmrTimersTimer.Start(); 

        }

        private void tmrTimersTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            //Call a delegate instantance on the UI thread using Invoke
            timeLabel.Invoke(new timerTick(this.UpdateGrid));
        }

        public void UpdateGrid(){
            //Update the time label
            timeLabel.Text = System.DateTime.Now.ToString("HH:mm:ss");
            dateLabel.Text = System.DateTime.Now.ToString("MM/dd/yyyy");
        }

    }
}
