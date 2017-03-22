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
    public partial class SS_Settings : Form
    {
        public SS_Settings()
        {
            InitializeComponent();
        }

        private void SS_Settings_Load(object sender, EventArgs e)
        {

            checkBox1.Checked = GlobalVars.SS0;
            checkBox2.Checked = GlobalVars.SS1;
            checkBox3.Checked = GlobalVars.SS2;
            checkBox4.Checked = GlobalVars.SS3;
            checkBox5.Checked = GlobalVars.SS4;
            checkBox6.Checked = GlobalVars.SS5;
            checkBox7.Checked = GlobalVars.SS6;
            checkBox8.Checked = GlobalVars.SS7;
            checkBox9.Checked = GlobalVars.SS8;
            checkBox10.Checked = GlobalVars.SS9;
            checkBox11.Checked = GlobalVars.SS10;
            checkBox12.Checked = GlobalVars.SS11;
            checkBox13.Checked = GlobalVars.SS12;
            checkBox14.Checked = GlobalVars.SS13;
            checkBox15.Checked = GlobalVars.SS14;
            checkBox16.Checked = GlobalVars.SS15;

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //save the values and then close
            GlobalVars.SS0 = checkBox1.Checked;
            GlobalVars.SS1 = checkBox2.Checked;
            GlobalVars.SS2 = checkBox3.Checked;
            GlobalVars.SS3 = checkBox4.Checked;
            GlobalVars.SS4 = checkBox5.Checked;
            GlobalVars.SS5 = checkBox6.Checked;
            GlobalVars.SS6 = checkBox7.Checked;
            GlobalVars.SS7 = checkBox8.Checked;
            GlobalVars.SS8 = checkBox9.Checked;
            GlobalVars.SS9 = checkBox10.Checked;
            GlobalVars.SS10 = checkBox11.Checked;
            GlobalVars.SS11 = checkBox12.Checked;
            GlobalVars.SS12 = checkBox13.Checked;
            GlobalVars.SS13 = checkBox14.Checked;
            GlobalVars.SS14 = checkBox15.Checked;
            GlobalVars.SS15 = checkBox16.Checked;

            this.Close();
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //save the values, but don't close
            GlobalVars.SS0 = checkBox1.Checked;
            GlobalVars.SS1 = checkBox2.Checked;
            GlobalVars.SS2 = checkBox3.Checked;
            GlobalVars.SS3 = checkBox4.Checked;
            GlobalVars.SS4 = checkBox5.Checked;
            GlobalVars.SS5 = checkBox6.Checked;
            GlobalVars.SS6 = checkBox7.Checked;
            GlobalVars.SS7 = checkBox8.Checked;
            GlobalVars.SS8 = checkBox9.Checked;
            GlobalVars.SS9 = checkBox10.Checked;
            GlobalVars.SS10 = checkBox11.Checked;
            GlobalVars.SS11 = checkBox12.Checked;
            GlobalVars.SS12 = checkBox13.Checked;
            GlobalVars.SS13 = checkBox14.Checked;
            GlobalVars.SS14 = checkBox15.Checked;
            GlobalVars.SS15 = checkBox16.Checked;
        }

        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false)
            {
                checkBox17.Checked = false;
            }
            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox7_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox7.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox9.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox10.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox11.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox12.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox13.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox14.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox15_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox15.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox16_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox16.Checked == false)
            {
                checkBox17.Checked = false;
            }

            if (checkBox1.Checked == true &&
                checkBox2.Checked == true &&
                checkBox3.Checked == true &&
                checkBox4.Checked == true &&
                checkBox5.Checked == true &&
                checkBox6.Checked == true &&
                checkBox7.Checked == true &&
                checkBox8.Checked == true &&
                checkBox9.Checked == true &&
                checkBox10.Checked == true &&
                checkBox11.Checked == true &&
                checkBox12.Checked == true &&
                checkBox13.Checked == true &&
                checkBox14.Checked == true &&
                checkBox15.Checked == true &&
                checkBox16.Checked == true)
            {
                checkBox17.Checked = true;
            }
        }

        private void checkBox17_Click(object sender, EventArgs e)
        {
            if (checkBox17.Checked == true)
            {
                checkBox1.Checked = true;
                checkBox2.Checked = true;
                checkBox3.Checked = true;
                checkBox4.Checked = true;
                checkBox5.Checked = true;
                checkBox6.Checked = true;
                checkBox7.Checked = true;
                checkBox8.Checked = true;
                checkBox9.Checked = true;
                checkBox10.Checked = true;
                checkBox11.Checked = true;
                checkBox12.Checked = true;
                checkBox13.Checked = true;
                checkBox14.Checked = true;
                checkBox15.Checked = true;
                checkBox16.Checked = true;
            }
            else
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
                checkBox15.Checked = false;
                checkBox16.Checked = false;
                checkBox17.Checked = false;
            }
        }
    }
}
