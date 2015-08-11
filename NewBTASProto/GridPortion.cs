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

namespace NewBTASProto
{

    public partial class Main_Form : Form
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
        public void SetUpTable()
        {

            // Add 16 rows to the data table to fit all of the grid data
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
                    updateD(i,a,temp[a]);
                }
            }
        }

        private void InitializeGrid()
        {
            #region Column Values
            // BTAS columns
            columnNames.Add("DT#");
            columnNames.Add("Work Order");
            columnNames.Add("Test");
            columnNames.Add("Step");
            columnNames.Add("In Use");
            columnNames.Add("Record");
            columnNames.Add("E-Time");
            columnNames.Add("Recording Status");
            columnNames.Add("Link Chgr");
            columnNames.Add("Chgr ID");
            columnNames.Add("Chgr Type");
            columnNames.Add("Chgr Status");

            // Create the columns using the columnNames list
            for (int i = 0; i < this.columnNames.Count; i++)
            {
                // The current process name.
                string name = this.columnNames[i];

                // Add the program name to our columns.
                if (name.ToString() == "In Use" || name.ToString() == "Record" || name.ToString() == "Link Chgr")
                {
                    d.Columns.Add(name, typeof(bool));
                }
                else
                {
                    d.Columns.Add(name, typeof(string));
                }

            }

            // Create the empty set of arrays to represent the 16 channels
            for (int i = 0; i < 16; i++)
            {
                channelArray.Add(new object[12] { i, "", "", "", 0, 0, "", "", 0, "", "", "" });
            }

            // Render the DataGridView.
            SetUpTable();
            dataGridView1.DataSource = d;

            // change settings for the individual columns
            for (int i = 0; i < 12; i++)
            {
                // these settings apply to every column
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].ReadOnly = true;
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // these setting only apply to specific columns
                switch (i)
                {
                    case 0:
                        dataGridView1.Columns[i].Width = 40;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.LightBlue;
                        break;
                    case 1:
                        dataGridView1.Columns[i].Width = 180;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.LightGoldenrodYellow;
                        break;
                    case 2:
                        dataGridView1.Columns[i].Width = 140;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.Aquamarine;
                        break;
                    case 3:
                        dataGridView1.Columns[i].Width = 40;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.Coral;
                        break;
                    case 4:
                        dataGridView1.Columns[i].Width = 44;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.Gainsboro;
                        break;
                    case 5:
                        dataGridView1.Columns[i].Width = 44;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.Gainsboro;
                        break;
                    case 6:
                        dataGridView1.Columns[i].Width = 100;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.LightCyan;
                        break;
                    case 7:
                        dataGridView1.Columns[i].Width = 120;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.Gainsboro;
                        break;
                    case 8:
                        dataGridView1.Columns[i].Width = 60;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.Gainsboro;
                        break;
                    case 9:
                        dataGridView1.Columns[i].Width = 50;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.LightGoldenrodYellow;
                        break;
                    case 10:
                        dataGridView1.Columns[i].Width = 78;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.LightGoldenrodYellow;
                        break;
                    case 11:
                        dataGridView1.Columns[i].Width = 78;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.Gainsboro;
                        break;
                }

            }

            //Now for a little row formatting
            for (int i = 0; i < 16; i++)
            {
                dataGridView1.Rows[i].Height = 20;
            }

            // finally, so we can clean up the startup jitter
            for (int j = 0; j < 16; j++)
            {
                dataGridView1.Rows[j].Cells[4].Style.BackColor = Color.Gainsboro;
            }
            #endregion


        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
            {
                return;
            }
            else if (e.ColumnIndex == 4 && (bool) d.Rows[e.RowIndex][5] != true)
            {
                if ((bool)d.Rows[e.RowIndex][e.ColumnIndex]) 
                {
                    updateD(e.RowIndex,e.ColumnIndex,false);
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Gainsboro;

                }
                else 
                {
                    updateD(e.RowIndex,e.ColumnIndex,true);
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Red;
                }
            }
            else if (e.ColumnIndex == 5)
            {
                if ((bool) d.Rows[e.RowIndex][e.ColumnIndex] == true)
                {
                    startNewTestToolStripMenuItem.Enabled = false;
                    stopTestToolStripMenuItem.Enabled = true;
                }
                else
                {
                    startNewTestToolStripMenuItem.Enabled = true;
                    stopTestToolStripMenuItem.Enabled = false;
                }
                cMSStartStop.Show(Cursor.Position);
            }
            else if (e.ColumnIndex == 8 && (bool)d.Rows[e.RowIndex][5] != true)
            {
                if ((bool)d.Rows[e.RowIndex][e.ColumnIndex])
                {
                    updateD(e.RowIndex,e.ColumnIndex,false);
                    updateD(e.RowIndex,11,"");
                    updateD(e.RowIndex,10,"");
                }
                else
                {
                    updateD(e.RowIndex,e.ColumnIndex,true);
                    if ((string) d.Rows[e.RowIndex][9] == "")
                    {
                        MessageBox.Show("You Still Need to Select a Charger ID Number");
                    }
                    else
                    {
                        checkForIC(int.Parse((string)d.Rows[e.RowIndex][9]),e.RowIndex);
                    }
                }
            }

            dataGridView1.ClearSelection();
        }



        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.ClearSelection();
            // only proceed if there isn't a test running!
            if ((bool)d.Rows[e.RowIndex][5] != true)
            {
                switch (e.ColumnIndex)
                {
                    case 1:
                        Choose_WO cwo = new Choose_WO(dataGridView1.CurrentRow.Index, (string)d.Rows[dataGridView1.CurrentRow.Index][1]);
                        cwo.Owner = this;
                        cwo.Show();
                        break;
                    case 2:
                        cMSTestType.Show(Cursor.Position);
                        break;
                    case 9:
                        cMSChargerChannel.Show(Cursor.Position);
                        break;
                    case 10:
                        cMSChargerType.Show(Cursor.Position);
                        break;
                }  // end switch
            }
            
        }

        //////////////////////////////////////////////////////////////////////locking stuff////////////////////////////
        private readonly object dLock = new object();

        private void updateD(int r, int c, object inVal)
        {
            lock (dLock)
            {

                d.Rows[r][c] = inVal;

            }
        }




    }
}