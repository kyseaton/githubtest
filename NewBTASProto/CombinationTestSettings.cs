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
            checkBox1.Checked = GlobalVars.FC6C1MinimumCellVotageAfterChargeTestEnabled;
            checkBox2.Checked = GlobalVars.DecliningCellVoltageTestEnabled;
            numericUpDown1.Value = GlobalVars.FC6C1MinimumCellVoltageThreshold;
            checkBox3.Checked = GlobalVars.FC6C1WaitEnabled;
            numericUpDown2.Value = GlobalVars.FC6C1WaitTime;
            checkBox5.Checked = GlobalVars.FC4C1MinimumCellVotageAfterChargeTestEnabled;
            numericUpDown4.Value = GlobalVars.FC4C1MinimumCellVoltageThreshold;
            checkBox4.Checked = GlobalVars.FC4C1WaitEnabled;
            numericUpDown3.Value = GlobalVars.FC4C1WaitTime;
            checkBox6.Checked = GlobalVars.CapTestVarEnable;
            numericUpDown5.Value = GlobalVars.CapTestVarValue;
            numericUpDown6.Value = GlobalVars.CSErr2Allow;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //save the values and then close
            GlobalVars.FC6C1MinimumCellVotageAfterChargeTestEnabled = checkBox1.Checked;
            GlobalVars.DecliningCellVoltageTestEnabled = checkBox2.Checked;
            GlobalVars.FC6C1MinimumCellVoltageThreshold = numericUpDown1.Value;
            GlobalVars.FC6C1WaitEnabled = checkBox3.Checked;
            GlobalVars.FC6C1WaitTime = numericUpDown2.Value;
            GlobalVars.FC4C1MinimumCellVotageAfterChargeTestEnabled = checkBox5.Checked;
            GlobalVars.FC4C1MinimumCellVoltageThreshold = numericUpDown4.Value;
            GlobalVars.FC4C1WaitEnabled = checkBox4.Checked;
            GlobalVars.FC4C1WaitTime = numericUpDown3.Value;
            GlobalVars.CapTestVarEnable = checkBox6.Checked;
            GlobalVars.CapTestVarValue = numericUpDown5.Value;
            GlobalVars.CSErr2Allow = numericUpDown6.Value;
            this.Close();
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }
    }
}
