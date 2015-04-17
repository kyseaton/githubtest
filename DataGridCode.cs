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

namespace WindowsFormsApplication1
{

    public partial class Form1 : Form
    {

        /// <summary>
        /// Contains column names
        /// </summary>
        List<string> columnNames = new List<string>();

        /// <summary>
        /// Contains column data arays.     
        /// </summary>
        List<object[]> channelArray = new List<object[]>();
        // Create the output table.
        DataTable d = new DataTable();

        /// <summary>
        /// This method builds the BTAS table
        /// </summary>
        public DataTable GetResultsTable()
        {

            // Add 16 rows to the data table to fit all of the channel data
            while (d.Rows.Count < 16)
            {
                d.Rows.Add();
            }

            //Now fill in the data table with data from each channel
            for (int i = 0; i < 16; i++)
            {

                object[] temp = this.channelArray[i];

                // Add each item to the cells in the column.
                for (int a = 0; a < 12; a++)
                {
                    d.Rows[i][a] = temp[a];
                }
            }
            return d;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

    }
}