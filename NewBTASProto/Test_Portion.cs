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
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;
using System.Text;

namespace NewBTASProto
{
    public partial class Main_Form : Form
        {

        public CancellationTokenSource cRunTest;

        private void RunTest()
        {
            cRunTest = new CancellationTokenSource();

            // Everything is going to be done on a helper thread
            ThreadPool.QueueUserWorkItem(s =>
            {
                
                // first we check if we have all the relavent options selected
                if ((string) d.Rows[dataGridView1.CurrentRow.Index][1] == "")
                {
                    MessageBox.Show("Please Assign a Work Order");
                    return;
                }
                else if ((string) d.Rows[dataGridView1.CurrentRow.Index][2] == "")
                {
                    MessageBox.Show("Please Select a Test Type");
                    return;
                }
                else if ((bool)d.Rows[dataGridView1.CurrentRow.Index][4] == false)
                {
                    MessageBox.Show("CScan is not In Use. Please Select it Before Proceeding");
                    return;
                }

                // Now we'll load the test parameters
                // We need to know the Interval and the number of readings



                // Now we'll look up the current test number and increment the new test

                // Save the test information to the test table
                
            },cRunTest.Token); // end thread

        }// end RunTest


    }
}
