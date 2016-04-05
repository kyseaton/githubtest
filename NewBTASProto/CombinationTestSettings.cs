using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NewBTASProto
{
    public partial class CombinationTestSettings : Form
    {
        public CombinationTestSettings()
        {
            InitializeComponent();
        }

        private void CombinationTestSettings_Load(object sender, EventArgs e)
        {
            checkBox1.Checked = Properties.Settings.Default.FC6C1MinimumCellVotageAfterChargeTestEnabled;
            checkBox2.Checked = Properties.Settings.Default.DecliningCellVoltageTestEnabled;
            numericUpDown1.Value = Properties.Settings.Default.FC6C1MinimumCellVoltageThreshold;
            checkBox3.Checked = Properties.Settings.Default.FC6C1WaitEnabled;
            numericUpDown2.Value = Properties.Settings.Default.FC6C1WaitTime;
            checkBox5.Checked = Properties.Settings.Default.FC4C1MinimumCellVotageAfterChargeTestEnabled;
            numericUpDown4.Value = Properties.Settings.Default.FC4C1MinimumCellVoltageThreshold;
            checkBox4.Checked = Properties.Settings.Default.FC4C1WaitEnabled;
            numericUpDown3.Value = Properties.Settings.Default.FC4C1WaitTime;
            checkBox6.Checked = Properties.Settings.Default.CapTestVarEnable;
            numericUpDown5.Value = Properties.Settings.Default.CapTestVarValue;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //save the values and then close
            Properties.Settings.Default.FC6C1MinimumCellVotageAfterChargeTestEnabled = checkBox1.Checked;
            Properties.Settings.Default.DecliningCellVoltageTestEnabled = checkBox2.Checked;
            Properties.Settings.Default.FC6C1MinimumCellVoltageThreshold = numericUpDown1.Value;
            Properties.Settings.Default.FC6C1WaitEnabled = checkBox3.Checked;
            Properties.Settings.Default.FC6C1WaitTime = numericUpDown2.Value;
            Properties.Settings.Default.FC4C1MinimumCellVotageAfterChargeTestEnabled = checkBox5.Checked;
            Properties.Settings.Default.FC4C1MinimumCellVoltageThreshold = numericUpDown4.Value;
            Properties.Settings.Default.FC4C1WaitEnabled = checkBox4.Checked;
            Properties.Settings.Default.FC4C1WaitTime = numericUpDown3.Value;
            Properties.Settings.Default.CapTestVarEnable = checkBox6.Checked;
            Properties.Settings.Default.CapTestVarValue = numericUpDown5.Value;
            Properties.Settings.Default.Save();
            this.Close();
        }
    }
}
