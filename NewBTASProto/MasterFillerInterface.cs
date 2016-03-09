using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Threading;

namespace NewBTASProto
{
    public partial class MasterFillerInterface : Form
    {
        //Now set the comboboxes to the current station and workorder...
        
        int curRow = 0;
        string curWorkOrder = "";
        float average = 0;
        int NumCells = 1;

        public MasterFillerInterface(int currentRow = 0, string currentWorkOrder = "",int NCells = 20)
        {
            InitializeComponent();
            loadWorkOrderList();

            curRow = currentRow;
            // we need to split up the work orders if we have multiple work orders on a single line...
            string tempWOS = currentWorkOrder;
            char[] delims = { ' ' };
            string[] A = tempWOS.Split(delims);
            curWorkOrder = A[0];
            if (NCells < 1)
            {
                NumCells = 20;
            }
            else
            {
                NumCells = NCells;
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void MasterFillerInterface_Load(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }
        private void loadWorkOrderList()
        {

            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT WorkOrderNumber FROM WorkOrders";

            DataSet workOrderList1 = new DataSet();
            OleDbConnection myAccessConn = null;
            // try to open the DB
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                lock (Main_Form.dataBaseLock)
                {
                    myAccessConn.Open();
                    myDataAdapter.Fill(workOrderList1, "ScanData");
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                
            }

            this.comboBox2.DisplayMember = "WorkOrderNumber";
            this.comboBox2.ValueMember = "WorkOrderNumber";
            this.comboBox2.DataSource = workOrderList1.Tables["ScanData"];
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            // this button house the routine to go and get the values from the masterfiller....
            GlobalVars.checkMasterFiller = true;
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;

            ThreadPool.QueueUserWorkItem(s =>
            {

                //loop until we get data....
                for (int i = 0; i < 20; i++)
                {
                    if (GlobalVars.checkMasterFiller == false) { break; }
                    Thread.Sleep(200);

                }// end for

                if (GlobalVars.checkMasterFiller == true)
                {
                    MessageBox.Show(this, "Failed to Read MasterFiller Data.  Please check your set up and try again.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    GlobalVars.checkMasterFiller = false;
                    this.Invoke((MethodInvoker)delegate
                    {
                        button1.Enabled = true;
                        button2.Enabled = true;
                        button3.Enabled = true;
                    });
                }
                else
                {

                    this.Invoke((MethodInvoker)delegate
                    {
                        numericUpDown1.Text = ((int.Parse(GlobalVars.MFData[5]) - 100) % 255).ToString();
                        numericUpDown2.Text = ((int.Parse(GlobalVars.MFData[6]) - 100) % 255).ToString();
                        numericUpDown3.Text = ((int.Parse(GlobalVars.MFData[7]) - 100) % 255).ToString();
                        numericUpDown4.Text = ((int.Parse(GlobalVars.MFData[8]) - 100) % 255).ToString();
                        numericUpDown5.Text = ((int.Parse(GlobalVars.MFData[9]) - 100) % 255).ToString();
                        numericUpDown6.Text = ((int.Parse(GlobalVars.MFData[10]) - 100) % 255).ToString();
                        numericUpDown7.Text = ((int.Parse(GlobalVars.MFData[11]) - 100) % 255).ToString();
                        numericUpDown8.Text = ((int.Parse(GlobalVars.MFData[12]) - 100) % 255).ToString();
                        numericUpDown9.Text = ((int.Parse(GlobalVars.MFData[13]) - 100) % 255).ToString();
                        numericUpDown10.Text = ((int.Parse(GlobalVars.MFData[14]) - 100) % 255).ToString();
                        numericUpDown11.Text = ((int.Parse(GlobalVars.MFData[15]) - 100) % 255).ToString();
                        numericUpDown12.Text = ((int.Parse(GlobalVars.MFData[16]) - 100) % 255).ToString();
                        numericUpDown13.Text = ((int.Parse(GlobalVars.MFData[17]) - 100) % 255).ToString();
                        numericUpDown14.Text = ((int.Parse(GlobalVars.MFData[18]) - 100) % 255).ToString();
                        numericUpDown15.Text = ((int.Parse(GlobalVars.MFData[19]) - 100) % 255).ToString();
                        numericUpDown16.Text = ((int.Parse(GlobalVars.MFData[20]) - 100) % 255).ToString();
                        numericUpDown17.Text = ((int.Parse(GlobalVars.MFData[21]) - 100) % 255).ToString();
                        numericUpDown18.Text = ((int.Parse(GlobalVars.MFData[22]) - 100) % 255).ToString();
                        numericUpDown19.Text = ((int.Parse(GlobalVars.MFData[23]) - 100) % 255).ToString();
                        numericUpDown20.Text = ((int.Parse(GlobalVars.MFData[24]) - 100) % 255).ToString();
                        numericUpDown21.Text = ((int.Parse(GlobalVars.MFData[25]) - 100) % 255).ToString();
                        numericUpDown22.Text = ((int.Parse(GlobalVars.MFData[26]) - 100) % 255).ToString();
                        numericUpDown23.Text = ((int.Parse(GlobalVars.MFData[27]) - 100) % 255).ToString();
                        numericUpDown24.Text = ((int.Parse(GlobalVars.MFData[28]) - 100) % 255).ToString();
                        // now fill in the average box...
                        average = 0;
                        int count = 0;
                        for (int i = 5; i < 28; i++)
                        {
                            if (((int.Parse(GlobalVars.MFData[i]) - 100) % 255) != 0)
                            {
                                count++;
                                average += (int.Parse(GlobalVars.MFData[i]) - 100) % 255;
                            }
                            
                        }
                        if(count == 0)
                        {
                            average = 0;                            
                        }
                        else
                        {
                            average /= count;
                        }
                        numericUpDown25.Value = (decimal) average;

                        button1.Enabled = true;
                        button2.Enabled = true;
                        button3.Enabled = true;
                    });
                }

            });

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //do we have a station?
            if (comboBox1.Text == "")
            {
                MessageBox.Show(this, "Please select a station to associate the MasterFiller data with", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //do we have a work order?
            if(comboBox2.Text == "")
            {
                MessageBox.Show(this, "Please select a Work Order to associate the MasterFiller data with", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // now fill recalc the average
            average = 0;
            average += (float) numericUpDown1.Value;        // always average the first cell...
            if (numericUpDown26.Value > 1) {
                average += (float)numericUpDown2.Value;
            }
            if (numericUpDown26.Value > 2)
            {
                average += (float)numericUpDown3.Value;
            }
            if (numericUpDown26.Value > 3)
            {
                average += (float)numericUpDown4.Value;
            }
            if (numericUpDown26.Value > 4)
            {
                average += (float)numericUpDown5.Value;
            }
            if (numericUpDown26.Value > 5)
            {
                average += (float)numericUpDown6.Value;
            }

            if (numericUpDown26.Value > 6)
            {
                average += (float)numericUpDown7.Value;
            }
            if (numericUpDown26.Value > 7)
            {
                average += (float)numericUpDown8.Value;
            }
            if (numericUpDown26.Value > 8)
            {
                average += (float)numericUpDown9.Value;
            }
            if (numericUpDown26.Value > 9)
            {
                average += (float)numericUpDown10.Value;
            }
            if (numericUpDown26.Value > 10)
            {
                average += (float)numericUpDown11.Value;
            }
            if (numericUpDown26.Value > 11)
            {
                average += (float)numericUpDown12.Value;
            }
            if (numericUpDown26.Value > 12)
            {
                average += (float)numericUpDown13.Value;
            }
            if (numericUpDown26.Value > 13)
            {
                average += (float)numericUpDown14.Value;
            }
            if (numericUpDown26.Value > 14)
            {
                average += (float)numericUpDown15.Value;
            }
            if (numericUpDown26.Value > 15)
            {
                average += (float)numericUpDown16.Value;
            }
            if (numericUpDown26.Value > 16)
            {
                average += (float)numericUpDown17.Value;
            }
            if (numericUpDown26.Value > 17)
            {
                average += (float)numericUpDown18.Value;
            }
            if (numericUpDown26.Value > 18)
            {
                average += (float)numericUpDown19.Value;
            }
            if (numericUpDown26.Value > 19)
            {
                average += (float)numericUpDown20.Value;
            }
            if (numericUpDown26.Value > 20)
            {
                average += (float)numericUpDown21.Value;
            }
            if (numericUpDown26.Value > 21)
            {
                average += (float)numericUpDown22.Value;
            }
            if (numericUpDown26.Value > 22)
            {
                average += (float)numericUpDown23.Value;
            }
            if (numericUpDown26.Value > 23)
            {
                average += (float)numericUpDown24.Value;
            }

            average /= (float) numericUpDown26.Value;


            numericUpDown25.Value = (decimal)average;

            //look up the workOrderID.../////////////////////////////////////////////////////////////////

            string strAccessConn;      
            OleDbConnection myAccessConn = null;

            // Open database containing all the battery data....
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            string strUpdateCMD = "INSERT INTO WaterLevel (WorkOrderNumber,Cell1,Cell2,Cell3,Cell4,Cell5,Cell6,Cell7,Cell8,Cell9,Cell10,Cell11,Cell12,Cell13,Cell14,Cell15,Cell16,Cell17,Cell18,Cell19,Cell20,Cell21,Cell22,Cell23,Cell24,AVE) "
                + "VALUES ('" +
                comboBox2.Text + "'," +                                                 //WorkOrderNumber
                numericUpDown1.Value.ToString("0") + "," +      //Cell1
                numericUpDown2.Value.ToString("0") + "," +      //Cell2
                numericUpDown3.Value.ToString("0") + "," +      //Cell3
                numericUpDown4.Value.ToString("0") + "," +      //Cell4
                numericUpDown5.Value.ToString("0") + "," +      //Cell5
                numericUpDown6.Value.ToString("0") + "," +     //Cell6
                numericUpDown7.Value.ToString("0") + "," +     //Cell7
                numericUpDown8.Value.ToString("0") + "," +     //Cell8
                numericUpDown9.Value.ToString("0") + "," +     //Cell9
                numericUpDown10.Value.ToString("0") + "," +     //Cell10
                numericUpDown11.Value.ToString("0") + "," +     //Cell11
                numericUpDown12.Value.ToString("0") + "," +     //Cell12
                numericUpDown13.Value.ToString("0") + "," +     //Cell13
                numericUpDown14.Value.ToString("0") + "," +     //Cell14
                numericUpDown15.Value.ToString("0") + "," +     //Cell15
                numericUpDown16.Value.ToString("0") + "," +     //Cell16
                numericUpDown17.Value.ToString("0") + "," +     //Cell17
                numericUpDown18.Value.ToString("0") + "," +     //Cell18
                numericUpDown19.Value.ToString("0") + "," +     //Cell19
                numericUpDown20.Value.ToString("0") + "," +     //Cell20
                numericUpDown21.Value.ToString("0") + "," +     //Cell21
                numericUpDown22.Value.ToString("0") + "," +     //Cell22
                numericUpDown23.Value.ToString("0") + "," +     //Cell23
                numericUpDown24.Value.ToString("0") + "," +     //Cell24
                average.ToString() +                                                    //Cell24
                ");";

            // try to open the DB
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to create a database connection. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);

                lock (Main_Form.dataBaseLock)
                {
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to store new data in the DataBase.\n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            finally
            {
                
            }
        }

        private void MasterFillerInterface_Shown(object sender, EventArgs e)
        {
            //Now set the comboboxes to the current station and workorder...
            comboBox1.Text = curRow.ToString();
            comboBox2.Text = curWorkOrder;
            numericUpDown26.Value = NumCells;
        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown26_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown25_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
