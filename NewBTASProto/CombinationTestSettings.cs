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

            checkBox2.Checked = GlobalVars.DecliningCellVoltageTestEnabled;
            checkBox6.Checked = GlobalVars.CapTestVarEnable;
            numericUpDown5.Value = GlobalVars.CapTestVarValue;
            numericUpDown6.Value = GlobalVars.CSErr2Allow;
            checkBox1.Checked = GlobalVars.showDeepDis;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //save the values and then close
            GlobalVars.DecliningCellVoltageTestEnabled = checkBox2.Checked;
            GlobalVars.CapTestVarEnable = checkBox6.Checked;
            GlobalVars.CapTestVarValue = numericUpDown5.Value;
            GlobalVars.CSErr2Allow = numericUpDown6.Value;
            GlobalVars.showDeepDis = checkBox1.Checked;
            this.Close();
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }
    }
}
