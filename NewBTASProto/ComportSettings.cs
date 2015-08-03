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

namespace NewBTASProto
{
    public partial class ComportSettings : Form
    {
        public ComportSettings()
        {
            InitializeComponent();
            // Get a list of serial port names. 
            string[] ports = SerialPort.GetPortNames();

            // Display each port name to the console. 
            foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
                comboBox2.Items.Add(port);
            }

            comboBox1.SelectedItem = GlobalVars.CSCANComPort;
            comboBox2.SelectedItem = GlobalVars.ICComPort;

            if (comboBox1.SelectedIndex == -1)
            {
                comboBox1.Items.Add(GlobalVars.CSCANComPort);
                comboBox1.SelectedItem = GlobalVars.CSCANComPort;
                label3.Visible = true;
            }

            if (comboBox2.SelectedIndex == -1)
            {
                comboBox2.Items.Add(GlobalVars.ICComPort);
                comboBox2.SelectedItem = GlobalVars.ICComPort;
                label4.Visible = true;
            }
            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // stop all of the scanning threads
            try
            {
                ((Main_Form)this.Owner).cPollIC.Cancel();
                ((Main_Form)this.Owner).cPollCScans.Cancel();
                ((Main_Form)this.Owner).sequentialScanT.Cancel();

                ((Main_Form)this.Owner).cPollIC.Dispose();
                ((Main_Form)this.Owner).cPollCScans.Dispose();
                ((Main_Form)this.Owner).sequentialScanT.Dispose();
            }
            catch (Exception ex)
            {
                if (ex is NullReferenceException || ex is ObjectDisposedException)
                {

                }
                else
                {
                    throw ex;
                }
            }
           

            // close the comms
            ((Main_Form)this.Owner).CSCANComPort.Close();
            ((Main_Form)this.Owner).ICComPort.Close();

            //Update the Globals
            GlobalVars.CSCANComPort = comboBox1.SelectedItem.ToString();
            GlobalVars.ICComPort = comboBox2.SelectedItem.ToString();

            //Make sure the warnings have been cleared
            ((Main_Form)this.Owner).label8.Visible = false;

            //Start the threads back up
            ((Main_Form)this.Owner).Scan();

            this.Dispose();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
}
