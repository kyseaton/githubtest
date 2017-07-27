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
using System.IO;
using System.Management;
using System.Threading;

namespace NewBTASProto
{
    public partial class ComportSettings : Form
    {
        bool comportSet = false;
        List<COMPortInfo> ComPortInformation = new List<COMPortInfo>();

        public ComportSettings()
        {
            InitializeComponent();
            SetupCOMPortInformation();
            // Get a list of serial port names. 
            string[] ports = SerialPort.GetPortNames();

            // Display each port name to the console. 
            int i = 0;
            foreach (COMPortInfo port in ComPortInformation)
            {
                if (port.friendlyName != null)
                {
                    comboBox1.Items.Add(port.friendlyName);
                }
                else
                {
                    comboBox1.Items.Add(port.portName);
                }
                if(port.portName == GlobalVars.CSCANComPort)
                {
                    comboBox1.SelectedIndex = i;
                }
                if(port.friendlyName != "")
                {
                    comboBox2.Items.Add(port.friendlyName);
                }
                else
                {
                    comboBox2.Items.Add(port.portName);
                }
                if(port.portName == GlobalVars.ICComPort)
                {
                    comboBox2.SelectedIndex = i;
                }
                i++;
            }

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

        private void SetupCOMPortInformation()
        {
            String[] portNames = System.IO.Ports.SerialPort.GetPortNames();
            foreach (String s in portNames)
            {
                // s is like "COM14"
                COMPortInfo ci = new COMPortInfo();
                ci.portName = s;
                ci.friendlyName = s;
                ComPortInformation.Add(ci);
            }

            String[] usbDevs = GetUSBCOMDevices();
            if (usbDevs == null)
            {
                //traditional
                return;
            }
            else
            {
                // with friendly
                foreach (String s in usbDevs)
                {
                    // Name will be like "USB Bridge (COM14)"
                    int start = s.IndexOf("(COM") + 1;
                    if (start >= 0)
                    {
                        int end = s.IndexOf(")", start + 3);
                        if (end >= 0)
                        {
                            // cname is like "COM14"
                            String cname = s.Substring(start, end - start);
                            for (int i = 0; i < ComPortInformation.Count; i++)
                            {
                                if (ComPortInformation[i].portName == cname)
                                {
                                    ComPortInformation[i].friendlyName = s;
                                }
                            }
                        }
                    }
                }
            }// end else
        }

        static string[] GetUSBCOMDevices()
        {
            try
            {
                return null;
                List<string> list = new List<string>();

                ManagementObjectSearcher searcher2 = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity");
                foreach (ManagementObject mo2 in searcher2.Get())
                {
                    string name = mo2["Name"].ToString();
                    // Name will have a substring like "(COM12)" in it.
                    if (name.Contains("(COM"))
                    {
                        list.Add(name);
                    }
                }
                // remove duplicates, sort alphabetically and convert to array
                string[] usbDevices = list.Distinct().OrderBy(s => s).ToArray();
                return usbDevices;
            }
            catch
            {
                return null;
            }
        }

        public class COMPortInfo
        {
            public String portName;
            public String friendlyName;
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
            ((Main_Form)this.Owner).CSCANComPort.Dispose();
            ((Main_Form)this.Owner).ICComPort.Close();
            ((Main_Form)this.Owner).ICComPort.Dispose();

            Thread.Sleep(1000);

            //Update the Globals
            int i = 0;
            foreach (COMPortInfo port in ComPortInformation)
            {
                if (port.friendlyName == comboBox1.Text)
                {
                    GlobalVars.CSCANComPort = port.portName;
                }
                if (port.friendlyName == comboBox2.Text)
                {
                    GlobalVars.ICComPort = port.portName;
                }
                i++;
            }

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
