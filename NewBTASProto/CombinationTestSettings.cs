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
            numericUpDown2.Value = GlobalVars.DecliningCellVoltageThres;
            checkBox6.Checked = GlobalVars.CapTestVarEnable;
            numericUpDown5.Value = GlobalVars.CapTestVarValue;
            numericUpDown6.Value = GlobalVars.CSErr2Allow;
            checkBox1.Checked = GlobalVars.showDeepDis;
            checkBox3.Checked = GlobalVars.allowZeroTest;
            numericUpDown1.Value = GlobalVars.rows2Dis;
            checkBox7.Checked = GlobalVars.advance2Short;
            checkBox4.Checked = GlobalVars.robustCSCAN;
            checkBox5.Checked = GlobalVars.InterpolateTime;
            checkBox8.Checked = GlobalVars.StopOnEnd;
            checkBox9.Checked = GlobalVars.AddOneMin;
            numericUpDown3.Value = GlobalVars.DCVPeriod;
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //save the values and then close
            GlobalVars.DecliningCellVoltageTestEnabled = checkBox2.Checked;
            GlobalVars.DecliningCellVoltageThres = numericUpDown2.Value;
            GlobalVars.InterpolateTime = checkBox5.Checked;
            GlobalVars.DCVPeriod = numericUpDown3.Value;
            GlobalVars.StopOnEnd = checkBox8.Checked;
            GlobalVars.AddOneMin = checkBox9.Checked;
            GlobalVars.CapTestVarEnable = checkBox6.Checked;
            GlobalVars.CapTestVarValue = numericUpDown5.Value;
            GlobalVars.CSErr2Allow = numericUpDown6.Value;
            GlobalVars.showDeepDis = checkBox1.Checked;
            GlobalVars.allowZeroTest = checkBox3.Checked;
            GlobalVars.rows2Dis = numericUpDown1.Value;
            GlobalVars.robustCSCAN = checkBox4.Checked;
            GlobalVars.advance2Short = checkBox7.Checked;
            ((Main_Form)this.Owner).dataGridView1.Height = Convert.ToInt32(27 + GlobalVars.rows2Dis * 20);
            ((Main_Form)this.Owner).groupBox3.Location = new Point(12, 422 - ((16 - Convert.ToInt32(GlobalVars.rows2Dis)) * 20));
            ((Main_Form)this.Owner).groupBox3.Height = ((Main_Form)this.Owner).Height - 485 + ((16 - Convert.ToInt32(GlobalVars.rows2Dis)) * 20);
            ((Main_Form)this.Owner).groupBox4.Location = new Point(((Main_Form)this.Owner).groupBox4.Location.X, 422 - ((16 - Convert.ToInt32(GlobalVars.rows2Dis)) * 20));
            ((Main_Form)this.Owner).groupBox4.Height = ((Main_Form)this.Owner).Height - 485 + ((16 - Convert.ToInt32(GlobalVars.rows2Dis)) * 20);
            this.Close();
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown5_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
