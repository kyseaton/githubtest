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
using System.Data.OleDb;

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
        public DataTable d = new DataTable();

        // Create graph settings table
        DataTable gs = new DataTable();

        // and the p(oint) c(olor) i(nformation) table
        public DataTable pci = new DataTable();


        /// <summary>
        /// This method builds the BTAS table
        /// </summary>
        public void SetUpTable()
        {

            d.TableName = "main_grid";

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
                for (int a = 0; a < 13; a++)
                {
                    updateD(i,a,temp[a]);
                }
            }
        }

        private void InitializeGrid()
        {

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
            columnNames.Add("Auto Config");

            // Create the columns using the columnNames list
            for (int i = 0; i < this.columnNames.Count; i++)
            {
                // The current process name.
                string name = this.columnNames[i];

                // Add the program name to our columns.
                if (name.ToString() == "In Use" || name.ToString() == "Record" || name.ToString() == "Link Chgr" || name.ToString() == "Auto Config")
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
                channelArray.Add(new object[13] { i, "", "", "", 0, 0, "", "", 0, "", "", "", 0 });
            }

            // Render the DataGridView.
            try
            {
                d.TableName = "main_grid";
                //System.IO.FileStream streamRead = new System.IO.FileStream(@"C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\main_grid.xml", System.IO.FileMode.Open, System.IO.FileAccess.Read);
                d.ReadXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\main_grid.xml");
                //streamRead.Close();

                if (d.Rows.Count == 1)
                {
                    SetUpTable();
                }
            }
            catch
            {
                SetUpTable();
                //error reading the grid!
            }
            
            dataGridView1.DataSource = d;

            // change settings for the individual columns
            for (int i = 0; i < 13; i++)
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
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.FloralWhite;
                        break;
                    case 2:
                        dataGridView1.Columns[i].Width = 140;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#ccffcc");
                        break;
                    case 3:
                        dataGridView1.Columns[i].Width = 40;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#ffbfa7");
                        break;
                    case 4:
                        dataGridView1.Columns[i].Width = 44;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.Gainsboro;
                        break;
                    case 5:
                        dataGridView1.Columns[i].Width = 44;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#ccffcc");
                        break;
                    case 6:
                        dataGridView1.Columns[i].Width = 100;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.LightCyan;
                        break;
                    case 7:
                        dataGridView1.Columns[i].Width = 120;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.LightCyan;
                        break;
                    case 8:
                        dataGridView1.Columns[i].Width = 60;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.Gainsboro;
                        break;
                    case 9:
                        dataGridView1.Columns[i].Width = 50;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#ccffcc");
                        break;
                    case 10:
                        dataGridView1.Columns[i].Width = 78;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.LightCyan;
                        break;
                    case 11:
                        dataGridView1.Columns[i].Width = 78;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.LightGoldenrodYellow;
                        break;
                    case 12:
                        dataGridView1.Columns[i].Width = 60;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.LightSkyBlue;
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

            //Also make sure that the linked chargers CSCANs know to hold them
            for (int j = 0; j < 16; j++)
            {
                if ((bool)d.Rows[j][8] == true)
                {
                    GlobalVars.cHold[j] = true;
                }
            }

            // so we do have a good data table.  Let's do some clean up on the DB to make sure we don't have any Active records what should be open
            // this is a problem after crashes...
            string strAccessConn;
            string strAccessSelect;
            OleDbConnection myAccessConn;

            // create the connection
            try
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // set all actives to open
            try
            {
                strAccessSelect = "UPDATE WorkOrders SET OrderStatus = 'Open' WHERE OrderStatus = 'Active'";
                OleDbCommand cmd = new OleDbCommand(strAccessSelect, myAccessConn);
                lock (Main_Form.dataBaseLock)
                {
                    myAccessConn.Open();
                    cmd.ExecuteNonQuery();
                    myAccessConn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Now we need to set the workorders that are actually active back to active
            // lets loop through it!

            try
            {
                for (int i = 0; i < 16; i++)
                {
                    if (d.Rows[i][1].ToString() != "")
                    {
                        strAccessSelect = "UPDATE WorkOrders SET OrderStatus = 'Active' WHERE WorkOrderNumber = '" + d.Rows[i][1].ToString() + "'";
                        OleDbCommand cmd = new OleDbCommand(strAccessSelect, myAccessConn);
                        lock (Main_Form.dataBaseLock)
                        {
                            myAccessConn.Open();
                            cmd.ExecuteNonQuery();
                            myAccessConn.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
            {
                return;
            }

            if (e.ColumnIndex == 4 && (bool) d.Rows[e.RowIndex][5] != true)
            {
                if ((bool)d.Rows[e.RowIndex][e.ColumnIndex]) 
                {
                    updateD(e.RowIndex,e.ColumnIndex,false);
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Gainsboro;
                    if (!d.Rows[e.RowIndex][9].ToString().Contains("S"))            // we don't have a slave...
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[8].Style.BackColor = Color.Gainsboro;
                        updateD(e.RowIndex, 10, "");
                        updateD(e.RowIndex, 11, "");
                    }

                    if (d.Rows[e.RowIndex][9].ToString().Contains("M"))            // we have a master!
                    {
                        //find the slave and clear it also..
                        string temp = d.Rows[e.RowIndex][9].ToString().Replace("-M", "");

                        for (int i = 0; i < 16; i++)
                        {
                            if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                            {
                                //found the slave
                                dataGridView1.Rows[i].Cells[8].Style.BackColor = Color.Gainsboro;
                                updateD(i, 10, "");
                                updateD(i, 11, "");
                                break;
                            }
                        }
                    }
                }
                else 
                {
                    updateD(e.RowIndex,e.ColumnIndex,true);
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Red;
                    if (d.Rows[e.RowIndex][9].ToString() != "" && !d.Rows[e.RowIndex][9].ToString().Contains("S"))
                    {
                        if (d.Rows[e.RowIndex][9].ToString().Length <= 2)
                        {
                            int chargerID = Convert.ToInt32(d.Rows[e.RowIndex][9]);
                            checkForIC(chargerID, e.RowIndex);
                        }
                        else if (d.Rows[e.RowIndex][9].ToString().Length == 3)
                        {
                            int chargerID = Convert.ToInt32(d.Rows[e.RowIndex][9].ToString().Substring(0, 1));
                            checkForIC(chargerID, e.RowIndex);
                        }
                        else
                        {
                            int chargerID = Convert.ToInt32(d.Rows[e.RowIndex][9].ToString().Substring(0, 2));
                            checkForIC(chargerID, e.RowIndex);
                        }
                    }
                }
            }
            else if (e.ColumnIndex == 5)
            {
                if (d.Rows[e.RowIndex][9].ToString().Contains("S"))
                {
                    // don't do anything with the slaves...
                    return;
                }

                if ((bool) d.Rows[e.RowIndex][e.ColumnIndex] == true)
                {
                    startNewTestToolStripMenuItem.Enabled = false;
                    resumeTestToolStripMenuItem.Enabled = false;
                    stopTestToolStripMenuItem.Enabled = true;
                }
                else
                {
                    startNewTestToolStripMenuItem.Enabled = true;
                    stopTestToolStripMenuItem.Enabled = false;
                    if ((string)d.Rows[e.RowIndex][6] == "")
                    {
                        resumeTestToolStripMenuItem.Enabled = false;
                    }
                    else
                    {
                        resumeTestToolStripMenuItem.Enabled = true;
                    }
                }
                cMSStartStop.Show(Cursor.Position);
            }
            else if (e.ColumnIndex == 8 && (bool)d.Rows[e.RowIndex][5] != true)
            {
                if (d.Rows[e.RowIndex][9].ToString().Contains("S"))
                {
                    // don't do anything with the slaves...
                    return;
                }
                else if ((bool)d.Rows[e.RowIndex][e.ColumnIndex])
                {
                    updateD(e.RowIndex,e.ColumnIndex,false);
                    updateD(e.RowIndex, 10, "");
                    updateD(e.RowIndex, 11, "");
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Gainsboro;
                    
                    // also update the slave (if we have a master...)
                    if (d.Rows[e.RowIndex][9].ToString().Contains("M"))
                    {
                        //find the slave
                        string temp = d.Rows[e.RowIndex][9].ToString().Replace("-M", "");

                        for (int i = 0; i < 16; i++)
                        {
                            if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                            {
                                //found the slave
                                updateD(i, 8, false);
                                updateD(i, 10, "");
                                updateD(i, 11, "");
                                dataGridView1.Rows[i].Cells[8].Style.BackColor = Color.LightSteelBlue;
                                break;
                            }
                        }
                    }

                    // Also clear the CSCAN to hold
                    GlobalVars.cHold[e.RowIndex] = false;

                }
                else
                {
                    updateD(e.RowIndex,e.ColumnIndex,true);
                    // also update the slave (if we have a master...)
                    if (d.Rows[e.RowIndex][9].ToString().Contains("M"))
                    {
                        //find the slave
                        string temp = d.Rows[e.RowIndex][9].ToString().Replace("-M", "");

                        for (int i = 0; i < 16; i++)
                        {
                            if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                            {
                                //found the slave
                                updateD(i, 8, true);
                                break;
                            }
                        }
                    }

                    if ((string) d.Rows[e.RowIndex][9] == "")
                    {
                        MessageBox.Show(this, "You Still Need to Select a Charger ID Number");
                    }
                    else
                    {
                        int chargerID = 0;

                        if (d.Rows[e.RowIndex][9].ToString().Length > 2)  // this is the case where we have a master and slave config
                        {
                            // we have a master slave charger
                            // split into 3 and 4 digit case
                            if (d.Rows[e.RowIndex][9].ToString().Length == 3)
                            {
                                if (d.Rows[e.RowIndex][9].ToString().Substring(2, 1) == "S") { return; }
                                // 3 case
                                chargerID = int.Parse(d.Rows[e.RowIndex][9].ToString().Substring(0, 1));
                            }
                            else
                            {
                                if (d.Rows[e.RowIndex][9].ToString().Substring(3, 1) == "S") { return; }
                                // 4 case
                                chargerID = int.Parse(d.Rows[e.RowIndex][9].ToString().Substring(0, 2));

                            }
                        }

                        else  // this is the normal case with just one charger
                        {
                            chargerID = Convert.ToInt32(d.Rows[e.RowIndex][9]);
                        }
                        checkForIC(chargerID,e.RowIndex);
                    }

                    // Also set the CSCAN to hold
                    GlobalVars.cHold[e.RowIndex] = true;
                }

            }
            else if (e.ColumnIndex == 12 && (bool)d.Rows[e.RowIndex][5] != true)
            {
                if (d.Rows[e.RowIndex][9].ToString().Contains("S"))
                {
                    //don't update the slave...
                    return;
                }
                else if (d.Rows[e.RowIndex][9].ToString().Contains("M"))
                {
                    //find the slave
                    string temp = d.Rows[e.RowIndex][9].ToString().Replace("-M", "");

                    for (int i = 0; i < 16; i++)
                    {
                        if (d.Rows[i][9].ToString().Contains(temp) && d.Rows[i][9].ToString().Contains("S"))
                        {
                            //found the slave
                            if ((bool)d.Rows[e.RowIndex][e.ColumnIndex])
                            {
                                updateD(e.RowIndex, e.ColumnIndex, false);
                                updateD(i, 12, false);
                            }
                            else
                            {
                                updateD(e.RowIndex, e.ColumnIndex, true);
                                updateD(i, 12, true);
                            }
                            break;
                        }
                    }
                }
                else
                {
                    //normal case...
                    if ((bool)d.Rows[e.RowIndex][e.ColumnIndex])
                    {
                        updateD(e.RowIndex, e.ColumnIndex, false);
                    }
                    else
                    {
                        updateD(e.RowIndex, e.ColumnIndex, true);
                    }
                }

            }

            dataGridView1.ClearSelection();
        }



        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {

                dataGridView1.ClearSelection();
                // only proceed if there isn't a test running!
                if ((bool)d.Rows[e.RowIndex][5] != true)
                {
                    switch (e.ColumnIndex)
                    {
                        case 1:
                            FormCollection fc = Application.OpenForms;
                            foreach (Form frm in fc)
                            {
                                if (frm is Choose_WO)
                                {
                                    if (frm.WindowState == FormWindowState.Minimized)
                                    {
                                        frm.WindowState = FormWindowState.Normal;
                                    }
                                    return;
                                }
                            }
                            Choose_WO cwo = new Choose_WO(dataGridView1.CurrentRow.Index, (string)d.Rows[dataGridView1.CurrentRow.Index][1]);
                            cwo.Owner = this;
                            cwo.Show();
                            break;
                        case 2:
                            if (d.Rows[e.RowIndex][9].ToString().Contains("S"))
                            {
                                // don't do anything with the slaves...
                                return;
                            }
                            cMSTestType.Show(Cursor.Position);
                            break;
                        case 9:
                            if ((bool)d.Rows[e.RowIndex][8] != true || (string)d.Rows[e.RowIndex][9] == "")
                            {
                                cMSChargerChannel.Show(Cursor.Position);
                            }
                            break;
                        case 10:
                            if ((bool)d.Rows[e.RowIndex][8] != true)
                            {
                                //cMSChargerType.Show(Cursor.Position);
                            }
                            break;
                    }  // end switch
                }
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

        bool startTog = true;

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (GlobalVars.loading == false)
            {
                try
                {
                    goodRead = false;
                    if (!startTog)
                    {
                        fillPlotCombos(e.RowIndex);
                    }
                    else
                    {
                        // put this in to not double up at startup on the fillPlotCombos
                        startTog = false;
                    }
                }
                // fill the plotCombos
                catch { ;}
            }

            
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {

        }


        private void Initialize_Graph_Settings()
        {
            try
            {
                //create the columns
                gs.Columns.Add("radio",typeof(bool));
                gs.Columns.Add("test",typeof(string));
                //name the table
                gs.TableName = "graph_set";
                //now read in what we got!
                gs.ReadXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\graph_set.xml");

                // do we have a good datatable?
                if (gs.Rows.Count != 16)
                {
                    gs.Clear();
                    for (int i = 0; i < 16; i++)
                    {
                        gs.Rows.Add();
                        gs.Rows[i][0] = false;
                        gs.Rows[i][1] = "Cell Voltages";
                    }
                    return;
                }

            }
            catch
            {
                //  we need to set the table up to be a null value table...
                gs.Clear();
                for (int i = 0; i < 16; i++)
                {
                    gs.Rows.Add();
                    gs.Rows[i][0] = false;
                    gs.Rows[i][1] = "Cell Voltages";
                }
            }
            
        }

        private void Initialize_PCI_Settings()
        {
            try
            {
                //create the columns
                pci.Columns.Add("bat", typeof(string));
                pci.Columns.Add("tech", typeof(string));
                pci.Columns.Add("NomV", typeof(float));
                pci.Columns.Add("NCells", typeof(int));
                pci.Columns.Add("BCVMIN", typeof(float));
                pci.Columns.Add("BCVMAX", typeof(float));
                pci.Columns.Add("CCVMMIN", typeof(float));
                pci.Columns.Add("CCVMAX", typeof(float));
                pci.Columns.Add("CCAPV", typeof(float));
                pci.Columns.Add("serial", typeof(string));

                //name the table
                pci.TableName = "pci_set";
                //now read in what we got!
                pci.ReadXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\pci_set.xml");

                // do we have a good datatable? Lets do some basic checks...
                if (pci.Rows.Count != 16 ||
                    pci.Rows[0][0].ToString() == "" ||
                    pci.Rows[0][1].ToString() == "" ||
                    pci.Rows[0][2].ToString() == "" ||
                    pci.Rows[0][3].ToString() == "" || 
                    pci.Rows[0][4].ToString() == "" ||
                    pci.Rows[0][5].ToString() == "" ||
                    pci.Rows[0][6].ToString() == "" ||
                    pci.Rows[0][7].ToString() == "" ||
                    pci.Rows[0][8].ToString() == "")
                {
                    pci.Clear();
                    for (int i = 0; i < 16; i++)
                    {
                        pci.Rows.Add();
                        pci.Rows[i][0] = "None";
                        pci.Rows[i][1] = "NiCd";
                        pci.Rows[i][2] = 24;         // negative 1 is the default...
                        pci.Rows[i][3] = -1;         // negative 1 is the default...
                        pci.Rows[i][4] = -1;         // negative 1 is the default...
                        pci.Rows[i][5] = -1;         // negative 1 is the default...
                        pci.Rows[i][6] = -1;         // negative 1 is the default...
                        pci.Rows[i][7] = 1.75;         // negative 1 is the default...
                        pci.Rows[i][8] = -1;         // negative 1 is the default...
                        pci.Rows[i][9] = "";
                    }
                }
            }
            catch
            {
                //  we need to set the table up to be a null value table...
                pci.Clear();
                for (int i = 0; i < 16; i++)
                {
                    pci.Rows.Add();
                    pci.Rows[i][0] = "None";
                    pci.Rows[i][1] = "NiCd";
                    pci.Rows[i][2] = 24;         // negative 1 is the default...
                    pci.Rows[i][3] = -1;         // negative 1 is the default...
                    pci.Rows[i][4] = -1;         // negative 1 is the default...
                    pci.Rows[i][5] = -1;         // negative 1 is the default...
                    pci.Rows[i][6] = -1;         // negative 1 is the default...
                    pci.Rows[i][7] = 1.75;         // negative 1 is the default...
                    pci.Rows[i][8] = -1;         // negative 1 is the default...
                    pci.Rows[i][9] = "";
                }
            }

        }

    }
}