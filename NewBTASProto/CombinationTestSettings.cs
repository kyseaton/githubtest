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
            Properties.Settings.Default.Save();
            this.Close();
        }
    }
}
