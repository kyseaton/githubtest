using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

namespace NewBTASProto
{

    public partial class ICSettingsForm : Form
    {
            Dictionary<string, int> testMode = new Dictionary<string, int>();
            Dictionary<int,string> reverseTestMode;
            Dictionary<string, int> action = new Dictionary<string, int>();
            Dictionary<int, string> reverseAction;

        public ICSettingsForm()
        {
            InitializeComponent();
            
            //fill up the dictionaries
            //
            //


            //testmode
            testMode.Add("10 Single Rate", 10);
            testMode.Add("11 Single Rate with Peak", 11);
            testMode.Add("12 Constant Voltage", 12);
            testMode.Add("20 Dual Rate", 20);
            testMode.Add("21 Dual Rate with Peak", 21);
            testMode.Add("30 Full Discharge", 30);
            testMode.Add("31 Capacity Test", 31);
            testMode.Add("32 Constant Resistance", 32);
            //now the reverse testmode
            reverseTestMode = testMode.ToDictionary(x => x.Value, x => x.Key);


            //action
            action.Add("Clear", 0);
            action.Add("Run", 1);
            action.Add("Stop", 2);
            action.Add("Reset", 3);

            //now the reverse action
            reverseAction = action.ToDictionary(x => x.Value, x => x.Key);

            // set up the numeric up down bounds
            // Charge Time 1
            numericUpDown1.Minimum = 0;             //hours
            numericUpDown1.Maximum = 99;
            numericUpDown2.Minimum = 0;             //mins
            numericUpDown2.Maximum = 59;
            numericUpDown3.Minimum = 0;             //charge current 1
            numericUpDown3.Maximum = 50;
            numericUpDown4.Minimum = 0;             //charge voltage 1
            numericUpDown4.Maximum = 77;

            // Charge Time 2
            numericUpDown8.Minimum = 0;             //hours
            numericUpDown8.Maximum = 99;
            numericUpDown7.Minimum = 0;             //mins
            numericUpDown7.Maximum = 59;
            numericUpDown6.Minimum = 0;             //charge current 2
            numericUpDown6.Maximum = 50;
            numericUpDown5.Minimum = 0;             //charge voltage 2
            numericUpDown5.Maximum = 77;

            // Discharge
            numericUpDown12.Minimum = 0;             //hours
            numericUpDown12.Maximum = 99;
            numericUpDown11.Minimum = 0;             //mins
            numericUpDown11.Maximum = 59;
            numericUpDown10.Minimum = 0;             //discharge current
            numericUpDown10.Maximum = 60;
            numericUpDown9.Minimum = 0;             //discharge voltage
            numericUpDown9.Maximum = 77;
            numericUpDown13.Minimum = 0;             //discharge resistance
            numericUpDown13.Maximum = 99;
        }

        private void ICSettingsForm_Shown(object sender, EventArgs e)
        {
            float remainder;

            try
            {
                comboBox1.SelectedIndex = ((Main_Form)this.Owner).dataGridView1.CurrentRow.Index;
                comboBox2.Text = reverseTestMode[GlobalVars.ICSettings[comboBox1.SelectedIndex].KM1 - 48];
                comboBox3.SelectedText = reverseAction[GlobalVars.ICSettings[comboBox1.SelectedIndex].KE3];
                // Primary Charge
                //time
                numericUpDown1.Value = GlobalVars.ICSettings[comboBox1.SelectedIndex].KM2 - 48;
                numericUpDown2.Value = GlobalVars.ICSettings[comboBox1.SelectedIndex].KM3 - 48;
                //current
                remainder = ((float)(GlobalVars.ICSettings[comboBox1.SelectedIndex].KM5 - 48) / 10);
                numericUpDown3.Value = (decimal) ((GlobalVars.ICSettings[comboBox1.SelectedIndex].KM4 - 48) * 10 + remainder);
                //voltage
                remainder = ((float)(GlobalVars.ICSettings[comboBox1.SelectedIndex].KM7- 48) / 100 );
                numericUpDown4.Value = (decimal) ((GlobalVars.ICSettings[comboBox1.SelectedIndex].KM6 - 48) + remainder);
                //Secondary Charge
                //time
                numericUpDown8.Value = GlobalVars.ICSettings[comboBox1.SelectedIndex].KM8 - 48;
                numericUpDown7.Value = GlobalVars.ICSettings[comboBox1.SelectedIndex].KM9 - 48;
                //current
                remainder = ((float)(GlobalVars.ICSettings[comboBox1.SelectedIndex].KM11 - 48) / 10);
                numericUpDown6.Value = (decimal)((GlobalVars.ICSettings[comboBox1.SelectedIndex].KM10 - 48) * 10 + remainder);
                //voltage
                remainder = ((float)(GlobalVars.ICSettings[comboBox1.SelectedIndex].KM13 - 48) / 100);
                numericUpDown5.Value = (decimal)((GlobalVars.ICSettings[comboBox1.SelectedIndex].KM12 - 48) + remainder);

                //Discharge
                //time
                numericUpDown12.Value = GlobalVars.ICSettings[comboBox1.SelectedIndex].KM14 - 48;
                numericUpDown11.Value = GlobalVars.ICSettings[comboBox1.SelectedIndex].KM15 - 48;
                //current
                remainder = ((float)(GlobalVars.ICSettings[comboBox1.SelectedIndex].KM17 - 48) / 10);
                numericUpDown10.Value = (decimal)((GlobalVars.ICSettings[comboBox1.SelectedIndex].KM16 - 48) * 10 + remainder);
                //voltage
                remainder = ((float)(GlobalVars.ICSettings[comboBox1.SelectedIndex].KM19 - 48) / 100);
                numericUpDown9.Value = (decimal)((GlobalVars.ICSettings[comboBox1.SelectedIndex].KM18 - 48) + remainder);
                //Ohms
                remainder = ((float)(GlobalVars.ICSettings[comboBox1.SelectedIndex].KM21 - 48) / 100);
                numericUpDown13.Value = (decimal)((GlobalVars.ICSettings[comboBox1.SelectedIndex].KM20 - 48) + remainder);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // set KE1 to data
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KE1 = (byte) 1;
            // update KM1
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM1 = (byte)(testMode[comboBox2.Text] + 48);

            // Charge Time 1
            //update KM2
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM2 = (byte)(numericUpDown1.Value + 48);
            //update KM3
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM3 = (byte)(numericUpDown2.Value + 48);

            // Charge Current 1
            // update KM4
             GlobalVars.ICSettings[comboBox1.SelectedIndex].KM4 = (byte)(numericUpDown3.Value / 10 + 48);
            //update KM5
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM5 = (byte)((numericUpDown3.Value % 10)*10 + 48);

            // Charge Voltage 1
            //update KM6
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM6 = (byte)(numericUpDown4.Value / 1 + 48);
            //update KM7
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM7 = (byte)((numericUpDown4.Value % 1)*100 + 48);

            ////////////////////////////////////////////////////////////////////////////////////////////////////

            // Charge Time 2
            //update KM8
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM8 = (byte)(numericUpDown8.Value + 48);
            //update KM9
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM9 = (byte)(numericUpDown7.Value + 48);

            // Charge Current 2
            // update KM10
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM10 = (byte)(numericUpDown6.Value / 10 + 48);
            //update KM11
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM11 = (byte)((numericUpDown6.Value % 10) * 10 + 48);

            // Charge Voltage 2
            //update KM12
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM12 = (byte)(numericUpDown5.Value / 1 + 48);
            //update KM13
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM13 = (byte)((numericUpDown5.Value % 1) * 100 + 48);

            ////////////////////////////////////////////////////////////////////////////////////////////////////

            // Discharge Time
            //update KM14
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM14 = (byte)(numericUpDown12.Value + 48);
            //update KM15
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM15 = (byte)(numericUpDown11.Value + 48);

            // Discharge Current
            // update KM16
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM16 = (byte)(numericUpDown10.Value / 10 + 48);
            //update KM17
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM17 = (byte)((numericUpDown10.Value % 10) * 10 + 48);

            // Discharge Voltage
            //update KM18
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM18 = (byte)(numericUpDown9.Value / 1 + 48);
            //update KM19
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM19 = (byte)((numericUpDown9.Value % 1) * 100 + 48);

            // Discharge Resistance
            //update KM20
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM20 = (byte)(numericUpDown13.Value / 1 + 48);
            //update KM21
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KM21 = (byte)((numericUpDown13.Value % 1) * 100 + 48);

            //Update the output string value
            GlobalVars.ICSettings[comboBox1.SelectedIndex].UpdateOutText();
            
            int inVal = comboBox1.SelectedIndex;
            //now we are going to create a thread to set KE1 back to data mode after 15 seconds
            ThreadPool.QueueUserWorkItem(s =>
            {
                Thread.Sleep(15000);
                // set KE1 to 0 ("data")
                GlobalVars.ICSettings[inVal].KE1 = (byte) 0;
                GlobalVars.ICSettings[inVal].UpdateOutText();
            }, inVal);                     // end thread

        }

        private void button2_Click(object sender, EventArgs e)
        {
            // set KE1 to 2 ("command")
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KE1 = (byte) 2;
            // reset KE3
            GlobalVars.ICSettings[comboBox1.SelectedIndex].KE3 = (byte)(action[comboBox3.Text]);
            //Update the output string value
            GlobalVars.ICSettings[comboBox1.SelectedIndex].UpdateOutText();

            int inVal = comboBox1.SelectedIndex;
            //now we are going to create a thread to set KE1 back to data mode after 15 seconds
            ThreadPool.QueueUserWorkItem(s =>
            {
                Thread.Sleep(15000);
                // set KE1 to 1 ("query")
                GlobalVars.ICSettings[inVal].KE1 = (byte) 0;
                GlobalVars.ICSettings[inVal].UpdateOutText();
            },inVal);                     // end thread

        }
    }
}
