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

        public MasterFillerInterface(int currentRow = 0, string currentWorkOrder = "")
        {
            InitializeComponent();
            loadWorkOrderList();

            curRow = currentRow;
            // we need to split up the work orders if we have multiple work orders on a single line...
            string tempWOS = currentWorkOrder;
            char[] delims = { ' ' };
            string[] A = tempWOS.Split(delims);
            curWorkOrder = A[0];
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
                    MessageBox.Show(this, "Failed to Read MasterFiller Data.  Please check your set up and try again.");
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
                        textBox1.Text = ((int.Parse(GlobalVars.MFData[5]) - 100) % 255).ToString();
                        textBox2.Text = ((int.Parse(GlobalVars.MFData[6]) - 100) % 255).ToString();
                        textBox3.Text = ((int.Parse(GlobalVars.MFData[7]) - 100) % 255).ToString();
                        textBox4.Text = ((int.Parse(GlobalVars.MFData[8]) - 100) % 255).ToString();
                        textBox5.Text = ((int.Parse(GlobalVars.MFData[9]) - 100) % 255).ToString();
                        textBox6.Text = ((int.Parse(GlobalVars.MFData[10]) - 100) % 255).ToString();
                        textBox7.Text = ((int.Parse(GlobalVars.MFData[11]) - 100) % 255).ToString();
                        textBox8.Text = ((int.Parse(GlobalVars.MFData[12]) - 100) % 255).ToString();
                        textBox9.Text = ((int.Parse(GlobalVars.MFData[13]) - 100) % 255).ToString();
                        textBox10.Text = ((int.Parse(GlobalVars.MFData[14]) - 100) % 255).ToString();
                        textBox11.Text = ((int.Parse(GlobalVars.MFData[15]) - 100) % 255).ToString();
                        textBox12.Text = ((int.Parse(GlobalVars.MFData[16]) - 100) % 255).ToString();
                        textBox13.Text = ((int.Parse(GlobalVars.MFData[17]) - 100) % 255).ToString();
                        textBox14.Text = ((int.Parse(GlobalVars.MFData[18]) - 100) % 255).ToString();
                        textBox15.Text = ((int.Parse(GlobalVars.MFData[19]) - 100) % 255).ToString();
                        textBox16.Text = ((int.Parse(GlobalVars.MFData[20]) - 100) % 255).ToString();
                        textBox17.Text = ((int.Parse(GlobalVars.MFData[21]) - 100) % 255).ToString();
                        textBox18.Text = ((int.Parse(GlobalVars.MFData[22]) - 100) % 255).ToString();
                        textBox19.Text = ((int.Parse(GlobalVars.MFData[23]) - 100) % 255).ToString();
                        textBox20.Text = ((int.Parse(GlobalVars.MFData[24]) - 100) % 255).ToString();
                        textBox21.Text = ((int.Parse(GlobalVars.MFData[25]) - 100) % 255).ToString();
                        textBox22.Text = ((int.Parse(GlobalVars.MFData[26]) - 100) % 255).ToString();
                        textBox23.Text = ((int.Parse(GlobalVars.MFData[27]) - 100) % 255).ToString();
                        textBox24.Text = ((int.Parse(GlobalVars.MFData[28]) - 100) % 255).ToString();
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
                        textBox25.Text = average.ToString();

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
                MessageBox.Show(this, "Please select a station to associate the MasterFiller data with");
                return;
            }

            //do we have a work order?
            if(comboBox2.Text == "")
            {
                MessageBox.Show(this, "Please select a Work Order to associate the MasterFiller data with");
                return;
            }
            //do we data?
            if (textBox1.Text == "")
            {
                MessageBox.Show(this, "Please aquire data before trying to save!");
                return;
            }

            //look up the workOrderID.../////////////////////////////////////////////////////////////////

            string strAccessConn;      
            OleDbConnection myAccessConn = null;

            // Open database containing all the battery data....
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            string strUpdateCMD = "INSERT INTO WaterLevel (WorkOrderNumber,Cell1,Cell2,Cell3,Cell4,Cell5,Cell6,Cell7,Cell8,Cell9,Cell10,Cell11,Cell12,Cell13,Cell14,Cell15,Cell16,Cell17,Cell18,Cell19,Cell20,Cell21,Cell22,Cell23,Cell24,AVE) "
                + "VALUES ('" +
                comboBox2.Text + "'," +                                                 //WorkOrderNumber
                ((int.Parse(GlobalVars.MFData[5]) - 100) % 255).ToString() + "," +      //Cell1
                ((int.Parse(GlobalVars.MFData[6]) - 100) % 255).ToString() + "," +      //Cell2
                ((int.Parse(GlobalVars.MFData[7]) - 100) % 255).ToString() + "," +      //Cell3
                ((int.Parse(GlobalVars.MFData[8]) - 100) % 255).ToString() + "," +      //Cell4
                ((int.Parse(GlobalVars.MFData[9]) - 100) % 255).ToString() + "," +      //Cell5
                ((int.Parse(GlobalVars.MFData[10]) - 100) % 255).ToString() + "," +     //Cell6
                ((int.Parse(GlobalVars.MFData[11]) - 100) % 255).ToString() + "," +     //Cell7
                ((int.Parse(GlobalVars.MFData[12]) - 100) % 255).ToString() + "," +     //Cell8
                ((int.Parse(GlobalVars.MFData[13]) - 100) % 255).ToString() + "," +     //Cell9
                ((int.Parse(GlobalVars.MFData[14]) - 100) % 255).ToString() + "," +     //Cell10
                ((int.Parse(GlobalVars.MFData[15]) - 100) % 255).ToString() + "," +     //Cell11
                ((int.Parse(GlobalVars.MFData[16]) - 100) % 255).ToString() + "," +     //Cell12
                ((int.Parse(GlobalVars.MFData[17]) - 100) % 255).ToString() + "," +     //Cell13
                ((int.Parse(GlobalVars.MFData[18]) - 100) % 255).ToString() + "," +     //Cell14
                ((int.Parse(GlobalVars.MFData[19]) - 100) % 255).ToString() + "," +     //Cell15
                ((int.Parse(GlobalVars.MFData[20]) - 100) % 255).ToString() + "," +     //Cell16
                ((int.Parse(GlobalVars.MFData[21]) - 100) % 255).ToString() + "," +     //Cell17
                ((int.Parse(GlobalVars.MFData[22]) - 100) % 255).ToString() + "," +     //Cell18
                ((int.Parse(GlobalVars.MFData[23]) - 100) % 255).ToString() + "," +     //Cell19
                ((int.Parse(GlobalVars.MFData[24]) - 100) % 255).ToString() + "," +     //Cell20
                ((int.Parse(GlobalVars.MFData[25]) - 100) % 255).ToString() + "," +     //Cell21
                ((int.Parse(GlobalVars.MFData[26]) - 100) % 255).ToString() + "," +     //Cell22
                ((int.Parse(GlobalVars.MFData[27]) - 100) % 255).ToString() + "," +     //Cell23
                ((int.Parse(GlobalVars.MFData[28]) - 100) % 255).ToString() + "," +     //Cell24
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
        }

        private void label27_Click(object sender, EventArgs e)
        {

        }
    }
}
