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
    public partial class Business_Name : Form
    {
        public Business_Name(object sender)
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.textBox1.Text = GlobalVars.businessName;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            GlobalVars.businessName = this.textBox1.Text;
            ((Main_Form)this.Owner).updateBusiness();
            this.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
