using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;

namespace NewBTASProto
{
    public partial class Main_Form : Form
    {
        public Main_Form()
        {
            InitializeComponent();
            Initialize_Menus_Tools();


            InitializeGrid();

            

            InitializeTimers();
            loadCommPorts();

        }

        //Replace with Dialog!
        private void loadCommPorts()
        {
            // Get the list of availible Ports
            string[] ArrayComPortsNames = null;
            int index = -1;
            string ComPortName = null;

            ArrayComPortsNames = SerialPort.GetPortNames();
            do
            {
                index += 1;
                cboPorts.Items.Add(ArrayComPortsNames[index]);
            }

            while (!((ArrayComPortsNames[index] == ComPortName)
                          || (index == ArrayComPortsNames.GetUpperBound(0))));
            Array.Sort(ArrayComPortsNames);

            //want to get first out
            if (index == ArrayComPortsNames.GetUpperBound(0))
            {
                ComPortName = ArrayComPortsNames[0];
            }
            cboPorts.Text = ArrayComPortsNames[ArrayComPortsNames.GetUpperBound(0)];

            // List the availible Baud Rates
            cboBaudRate.Items.Add(300);
            cboBaudRate.Items.Add(600);
            cboBaudRate.Items.Add(1200);
            cboBaudRate.Items.Add(2400);
            cboBaudRate.Items.Add(4800);
            cboBaudRate.Items.Add(9600);
            cboBaudRate.Items.Add(14400);
            cboBaudRate.Items.Add(19200);
            cboBaudRate.Items.Add(38400);
            cboBaudRate.Items.Add(57600);
            cboBaudRate.Items.Add(115200);
            cboBaudRate.Text = cboBaudRate.Items[7].ToString();

            //List Data Bit Settings
            cboDataBits.Items.Add(7);
            cboDataBits.Items.Add(8);
            cboDataBits.Text = cboDataBits.Items[1].ToString();

            //List Stop Bit Options
            cboStopBits.Items.Add("One");
            cboStopBits.Items.Add("OnePointFive");
            cboStopBits.Items.Add("Two");
            cboStopBits.Text = cboStopBits.Items[0].ToString();

            // List Parity Settings
            cboParity.Items.Add("None");
            cboParity.Items.Add("Even");
            cboParity.Items.Add("Mark");
            cboParity.Items.Add("Odd");
            cboParity.Items.Add("Space");
            cboParity.Text = cboParity.Items[0].ToString();

            //List Handshake Settings
            cboHandShaking.Items.Add("None");
            cboHandShaking.Items.Add("XOnXOff");
            cboHandShaking.Items.Add("RequestToSend");
            cboHandShaking.Items.Add("RequestToSendXOnXOff");
            cboHandShaking.Text = cboHandShaking.Items[0].ToString();
        }

        /// <summary>
        /// This function will update all of the Menus with the appropriate values
        /// </summary>
        private void Initialize_Menus_Tools()
        {
            if (GlobalVars.useF) { this.fahrenheitToolStripMenuItem.Checked = true; }
            else{this.centigradeToolStripMenuItem.Checked = true;}

            if (GlobalVars.Pos2Neg) { this.positiveToNegativeToolStripMenuItem.Checked = true; }
            else { this.negativeToPositiveToolStripMenuItem.Checked = true; }

            toolStripStatusLabel1.Text = "Version:  " + Application.ProductVersion;

            label10.Text = GlobalVars.businessName;

        }


        // REMOVE!
        private void btnPortState_Click(object sender, EventArgs e)
        {
            if (btnPortState.Text == "Closed")
            {
                sendCScan.Enabled = true;
                btnPortState.Text = "Open";
                ComPort.PortName = Convert.ToString(cboPorts.Text);
                ComPort.BaudRate = Convert.ToInt32(cboBaudRate.Text);
                ComPort.DataBits = Convert.ToInt32(cboDataBits.Text);
                ComPort.StopBits = (StopBits)Enum.Parse(typeof(StopBits), cboStopBits.Text);
                ComPort.Handshake = (Handshake)Enum.Parse(typeof(Handshake), cboHandShaking.Text);
                ComPort.Parity = (Parity)Enum.Parse(typeof(Parity), cboParity.Text);
                ComPort.Open();
            }
            else if (btnPortState.Text == "Open")
            {
                sendCScan.Enabled = false;
                btnPortState.Text = "Closed";
                ComPort.Close();
            }
        }

        // REMOVE!
        private void sendCScan_Click(object sender, EventArgs e)
        {
            string outText = ("~" + tbT1.Text.ToString() + tbT2.Text.ToString() + tbWDO.Text.ToString() + tbKM1.Text.ToString() + tbKM2.Text.ToString() + "Z");
            ComPort.Write(outText);
        }

        private void toolStripStatusLabel3_Click(object sender, EventArgs e)
        {

        }

        private void test3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Graphics_Form gf = new Graphics_Form();
            gf.Show();
   
        }

        private void testToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Reports_Form rf = new Reports_Form();
            rf.Show();
        }

        private void Main_Form_FormClosed(object sender, FormClosedEventArgs e)
        {
            SaveGlobals();
            Application.Exit();
        }

        private void SaveGlobals()
        {

            string strAccessConn;
            string strUpdateCMD;
            // Open database containing all the battery data....
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
            strUpdateCMD = "UPDATE Options SET Degree='" + (GlobalVars.useF ? "F." : "C.") + "', CellOrder='" + (GlobalVars.Pos2Neg ? "Pos. to Neg." : "Neg. to Pos.") + "', BusinessName='"+ GlobalVars.businessName+"';";
            OleDbConnection myAccessConn;

            // try to open the DB
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                myAccessConn.Open();
                myAccessCommand.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                return;
            }
            finally
            {
                myAccessConn.Close();
            }

        }

        private void bussinessNameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Business_Name bn = new Business_Name(this);
            bn.Owner = this;
            bn.Show();
        }

        private void centigradeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.centigradeToolStripMenuItem.Checked == false) 
            { 
                this.centigradeToolStripMenuItem.Checked = true;
                this.fahrenheitToolStripMenuItem.Checked = false;
                GlobalVars.useF = false;
            }
        }

        private void fahrenheitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.fahrenheitToolStripMenuItem.Checked == false)
            {
                this.fahrenheitToolStripMenuItem.Checked = true;
                this.centigradeToolStripMenuItem.Checked = false;
                GlobalVars.useF = true;
            }

        }

        private void negativeToPositiveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.negativeToPositiveToolStripMenuItem.Checked == false)
            {
                this.negativeToPositiveToolStripMenuItem.Checked = true;
                this.positiveToNegativeToolStripMenuItem.Checked = false;
                GlobalVars.Pos2Neg = false;
            }
        }

        private void positiveToNegativeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.positiveToNegativeToolStripMenuItem.Checked == false)
            {
                this.positiveToNegativeToolStripMenuItem.Checked = true;
                this.negativeToPositiveToolStripMenuItem.Checked = false;
                GlobalVars.Pos2Neg = true;
            }
        }

        private void newCustomBatteryToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        public void updateBusiness() 
        {
            label10.Text = GlobalVars.businessName;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            chart1.Series.Clear();
            button5.Enabled = false;
            pollCScan(1);

        }

    }
}
