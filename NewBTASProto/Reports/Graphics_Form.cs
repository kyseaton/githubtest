using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Xml.Serialization;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Threading;

namespace NewBTASProto
{
    public partial class Graphics_Form : Form
    {
        //form wide dataSets used to fill in the two graphs
        DataSet graph1Set = new DataSet();
        DataSet graph2Set = new DataSet();

        DataTable customTestParams;

        //When the type of battery the technology is retrieved store it here
        string technology1 = "NiCd";
        string technology2 = "NiCd";

        string type1 = "";
        string type2 = "";

        int cell1 = 0;
        int cell2 = 0;

        Title tt1 = new Title();
        Title tt2 = new Title();

        string batMod1;
        string batMod2;

        float NomV1;
        float NomV2;

        float CellV1;
        float CellV2;

        string curStep = "";
        string curWorkOrder = "";
        bool startup = true;
        bool startup2 = true;

        DataSet batInfo1;
        DataSet batInfo2;

        public Graphics_Form(string currentStep = "", string currentWorkOrder = "")
        {
            InitializeComponent();

            curStep = currentStep;
            curWorkOrder = currentWorkOrder;
        }

        private void Graphics_Form_Load(object sender, EventArgs e)
        {
            customTestParams = ((Main_Form)this.Owner).customTestParams;

            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture;
            loadWorkOrderLists();

            // Need to set up the title for the charts...
            tt1.Name = "tTitle";
            tt1.Text = "";
            chart1.Titles.Add(tt1);
            tt2.Name = "tTitle";
            tt2.Text = "";
            chart2.Titles.Add(tt2);

            checkBox1.Checked = GlobalVars.dualPlots;

            //myDataSet.Tables[0].

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                label2.Visible = true;
                comboBox3.Visible = true;
                comboBox4.Visible = true;
                radioButton3.Visible = true;
                radioButton4.Visible = true;
                chart2.Visible = true;
                label1.Visible = true;
                comboBox7.Visible = true;
                comboBox8.Visible = true;
                groupBox2.Visible = true;
                
                chart1.Height = (panel1.Height / 2) - 3;
                chart2.Height = (panel1.Height / 2) - 3;
            }
            else
            {
                label2.Visible = false;
                comboBox3.Visible = false;
                comboBox4.Visible = false;
                radioButton3.Visible = false;
                radioButton4.Visible = false;
                chart2.Visible = false;
                label1.Visible = false;
                comboBox7.Visible = false;
                comboBox8.Visible = false;
                groupBox2.Visible = false;
                chart2.Visible = false;
                chart1.Height = panel1.Height - 6;
            }
        }

        private void loadWorkOrderLists()
        {
            comboBox1.Enabled = false;
            comboBox1.Text = "Loading";
            comboBox3.Enabled = false;
            comboBox3.Text = "Loading";

            // do this on a helper thread!
            ThreadPool.QueueUserWorkItem(s =>
            {

                string strAccessConn;
                string strAccessSelect;
                // Open database containing all the battery data....
                if (checkBox2.Checked == true)
                {
                    strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                    strAccessSelect = @"SELECT WorkOrderNumber FROM WorkOrders";
                }
                else
                {
                    strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                    strAccessSelect = @"SELECT WorkOrderNumber FROM WorkOrders WHERE OrderStatus<>'Archived'";
                }


            
                DataSet workOrderList1 = new DataSet();
                DataSet workOrderList2 = new DataSet();
                OleDbConnection myAccessConn = null;
                // try to open the DB
                try
                {
                    myAccessConn = new OleDbConnection(strAccessConn);
                }
                catch (Exception ex)
                {
                    this.Invoke((MethodInvoker)delegate()
                    {
                        MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
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
                        myDataAdapter.Fill(workOrderList2, "ScanData");
                        myAccessConn.Close();
                    }



                }
                catch (Exception ex)
                {
                    myAccessConn.Close();
                    this.Invoke((MethodInvoker)delegate()
                    {
                        MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }); 
                    return;
                }

                DataRow emptyRow1 = workOrderList1.Tables["ScanData"].NewRow();
                emptyRow1["WorkOrderNumber"] = "";
                workOrderList1.Tables["ScanData"].Rows.InsertAt(emptyRow1,0);

                DataRow emptyRow2 = workOrderList2.Tables["ScanData"].NewRow();
                emptyRow2["WorkOrderNumber"] = "";
                workOrderList2.Tables["ScanData"].Rows.InsertAt(emptyRow2, 0);

                this.Invoke((MethodInvoker)delegate()
                {                
                    this.comboBox1.DisplayMember = "WorkOrderNumber";
                    this.comboBox1.ValueMember = "WorkOrderNumber";
                    this.comboBox1.DataSource = workOrderList1.Tables["ScanData"];

                    this.comboBox3.DisplayMember = "WorkOrderNumber";
                    this.comboBox3.ValueMember = "WorkOrderNumber";
                    this.comboBox3.DataSource = workOrderList2.Tables["ScanData"];

                    // remember to clear everything!
                    this.comboBox2.DataSource = null;
                    this.comboBox4.DataSource = null;
                    tt1.Text = "";
                    tt2.Text = "";

                    comboBox1.Enabled = true;
                    comboBox3.Enabled = true;

                    if (startup)
                    {
                        //Now set the comboboxes to the current station and workorder...

                        // we need to split up the work orders if we have multiple work orders on a single line...
                        string tempWOS = curWorkOrder;
                        char[] delims = { ' ' };
                        string[] A = tempWOS.Split(delims);
                        curWorkOrder = A[0];

                        comboBox1.Text = curWorkOrder.Trim();
                        comboBox3.Text = curWorkOrder.Trim();
                        //comboBox1_SelectedValueChanged(this, null);

                        if (curStep == "")
                        {
                            startup = false;
                            startup2 = false;
                        }
                    }
                });

            }); // end thread
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            chart1.Series.Clear();
            tt1.Text = "";
            string combo1Text = comboBox1.Text;


            if (comboBox1.SelectedIndex <= 0) { return; }

            else
            {
                comboBox2.Enabled = false;
                comboBox2.Text = "Loading";

                // do it on a helper thread!
                ThreadPool.QueueUserWorkItem(s =>
                {
                    // reset cells
                    cell1 = 0;

                    // first clear the graph type comboboxes
                    this.Invoke((MethodInvoker)delegate()
                    {
                        comboBox5.Items.Clear();
                        comboBox5.Text = "";
                        comboBox6.Items.Clear();
                        comboBox6.Text = "";
                    });

                    // Open database containing all the battery data....
                    string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                    string strAccessSelect = @"SELECT StepNumber,TestName, Technology, CustomNoCells FROM Tests WHERE WorkOrderNumber='" + combo1Text + @"' ORDER BY StepNumber ASC";


                    DataSet testsPerformed = new DataSet();
                    OleDbConnection myAccessConn = null;
                    // try to open the DB
                    try
                    {
                        myAccessConn = new OleDbConnection(strAccessConn);
                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }); 
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
                            myDataAdapter.Fill(testsPerformed, "Tests");
                            myAccessConn.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }); 
                        return;
                    }

                    try
                    {
                        DataRow emptyRow1 = testsPerformed.Tables["Tests"].NewRow();
                        emptyRow1["TestName"] = "";
                        testsPerformed.Tables["Tests"].Rows.InsertAt(emptyRow1, 0);
                        testsPerformed.Tables["Tests"].Columns.Add("ForList", typeof(string), "StepNumber + ' - '+ TestName");

                        this.Invoke((MethodInvoker)delegate()
                        {
                            this.comboBox2.DisplayMember = "ForList";
                            this.comboBox2.ValueMember = "StepNumber";
                            this.comboBox2.DataSource = testsPerformed.Tables["Tests"];

                            comboBox2.Enabled = true;

                            if (startup)
                            {
                                try
                                {
                                    comboBox2.SelectedIndex = comboBox2.FindString(curStep);
                                }
                                catch
                                {
                                    // do nothing...
                                }
                            }
                        });

                        // update the technology var
                        //technology1 = testsPerformed.Tables["Tests"].Rows[1][2].ToString();
                        if (testsPerformed.Tables["Tests"].Rows[1][3].ToString() != "") { cell1 = Int32.Parse(testsPerformed.Tables["Tests"].Rows[1][3].ToString()); }
                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        return;
                    }

                });

            }
        }
        private void comboBox3_SelectedValueChanged(object sender, EventArgs e)
        {
            chart2.Series.Clear();
            tt2.Text = "";
            string combo3Text = comboBox3.Text;


            if (comboBox3.SelectedIndex <= 0) { return; }
            else
            {
                comboBox4.Enabled = false;
                comboBox4.Text = "Loading";
                // do it on a helper thread!
                ThreadPool.QueueUserWorkItem(s =>
                {
                    // reset cells
                    cell2 = 0;

                    // first clear the graph type comboboxes
                    this.Invoke((MethodInvoker)delegate()
                    {
                        comboBox7.Items.Clear();
                        comboBox7.Text = "";
                        comboBox8.Items.Clear();
                        comboBox8.Text = "";
                    });

                    // Open database containing all the battery data....
                    string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                    string strAccessSelect = @"SELECT StepNumber,TestName, Technology, CustomNoCells FROM Tests WHERE WorkOrderNumber='" + combo3Text + @"' ORDER BY StepNumber ASC";


                    DataSet testsPerformed = new DataSet();
                    OleDbConnection myAccessConn = null;
                    // try to open the DB
                    try
                    {
                        myAccessConn = new OleDbConnection(strAccessConn);
                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }); return;
                    }
                    //  now try to access it
                    try
                    {
                        OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                        lock (Main_Form.dataBaseLock)
                        {
                            myAccessConn.Open();
                            myDataAdapter.Fill(testsPerformed, "Tests");
                            myAccessConn.Close();
                        }


                    }
                    catch (Exception ex)
                    {   
                        myAccessConn.Close();
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        return;
                    }

                    try
                    {
                        DataRow emptyRow1 = testsPerformed.Tables["Tests"].NewRow();
                        emptyRow1["TestName"] = "";
                        testsPerformed.Tables["Tests"].Rows.InsertAt(emptyRow1, 0);
                        testsPerformed.Tables["Tests"].Columns.Add("ForList", typeof(string), "StepNumber + ' - '+ TestName");

                        this.Invoke((MethodInvoker)delegate()
                        {
                            this.comboBox4.DisplayMember = "ForList";
                            this.comboBox4.ValueMember = "StepNumber";
                            this.comboBox4.DataSource = testsPerformed.Tables["Tests"];

                            comboBox4.Enabled = true;

                            if (startup2)
                            {
                                try
                                {
                                    comboBox4.SelectedIndex = comboBox4.FindString(curStep);
                                }
                                catch
                                {
                                    // do nothing...
                                }
                            }
                        });

                        // update the technology var
                        //technology2 = testsPerformed.Tables["Tests"].Rows[1][2].ToString();
                        //if (Int32.Parse(testsPerformed.Tables["Tests"].Rows[1][3].ToString()) != 0) { cell2 = Int32.Parse(testsPerformed.Tables["Tests"].Rows[1][3].ToString()); }
                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        return;
                    }
                });

            }
        }


        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {

            chart1.Series.Clear();
            tt1.Text = "";

            string combo1Text = comboBox1.Text;
            string combo2Text = comboBox2.Text;
            if (comboBox2.SelectedIndex <= 0) { return; }
            else
            {
                comboBox5.Enabled = false;
                comboBox5.Text = "Loading";
                comboBox6.Enabled = false;
                comboBox6.Text = "Loading";

                // do it on a helper thread!
                ThreadPool.QueueUserWorkItem(s =>
                {
                    // FIRST CLEAR THE OLD DATA SET!
                    graph1Set.Clear();
                    // Open database containing all the battery data....
                    string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                    string strAccessSelect = @"SELECT * FROM ScanData WHERE BWO='" + combo1Text + @"' AND STEP='" + combo2Text.Substring(0, 2) + @"' ORDER BY RDG ASC";

                    //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
                    OleDbConnection myAccessConn = null;
                    // try to open the DB
                    try
                    {
                        myAccessConn = new OleDbConnection(strAccessConn);
                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
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
                            myDataAdapter.Fill(graph1Set, "ScanData");
                            myAccessConn.Close();
                        }


                    }
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        return;
                    }

                    // we have the data table at this point, but we still need to update the graph combo boxes
                    // first we look up the cable used for the test

                    // Open database containing all the battery data....
                    strAccessSelect = @"SELECT CellCableID FROM Tests WHERE WorkOrderNumber='" + combo1Text + @"' AND StepNumber='" + combo2Text.Substring(0, 2) + @"'";


                    DataSet lookForCable = new DataSet();
                    myAccessConn = null;
                    // try to open the DB
                    try
                    {
                        myAccessConn = new OleDbConnection(strAccessConn);
                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
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
                            myDataAdapter.Fill(lookForCable, "Tests");
                            myAccessConn.Close();
                        }


                    }
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        return;
                    }


                    string cellCable = lookForCable.Tables["Tests"].Rows[0][0].ToString();

                    this.Invoke((MethodInvoker)delegate()
                    {
                        // don't let the user look at cell values for SLA bats...
                        if (cellCable != "10")
                        {
                            radioButton2.Enabled = true;
                        }
                        else
                        {
                            radioButton2.Enabled = false;
                            radioButton1.Checked = true;
                        }

                        switch (cellCable)
                        {
                            case "1":
                            case "23":
                                // update the cells value
                                if (cell1 == 0) { cell1 = 20; }
                                // Battery combobox
                                comboBox5.Items.Clear();
                                //comboBox5.Text = "";
                                comboBox5.Items.Add("Voltage");
                                comboBox5.Items.Add("Current");
                                comboBox5.Items.Add("Temperature 1");
                                comboBox5.Items.Add("Temperature 2");
                                comboBox5.Items.Add("Temperature 3");
                                comboBox5.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox6.Items.Clear();
                                //comboBox6.Text = "";
                                comboBox6.Items.Add("Ending Voltages");
                                comboBox6.Items.Add("Cell 1");
                                comboBox6.Items.Add("Cell 2");
                                comboBox6.Items.Add("Cell 3");
                                comboBox6.Items.Add("Cell 4");
                                comboBox6.Items.Add("Cell 5");
                                comboBox6.Items.Add("Cell 6");
                                comboBox6.Items.Add("Cell 7");
                                comboBox6.Items.Add("Cell 8");
                                comboBox6.Items.Add("Cell 9");
                                comboBox6.Items.Add("Cell 10");
                                comboBox6.Items.Add("Cell 11");
                                comboBox6.Items.Add("Cell 12");
                                comboBox6.Items.Add("Cell 13");
                                comboBox6.Items.Add("Cell 14");
                                comboBox6.Items.Add("Cell 15");
                                comboBox6.Items.Add("Cell 16");
                                comboBox6.Items.Add("Cell 17");
                                comboBox6.Items.Add("Cell 18");
                                comboBox6.Items.Add("Cell 19");
                                comboBox6.Items.Add("Cell 20");
                                break;
                            case "2":
                                // update the cells value
                                if (cell1 == 0) { cell1 = 20; }
                                // Battery combobox
                                comboBox5.Items.Clear();
                                //comboBox5.Text = "";
                                comboBox5.Items.Add("Voltage");
                                comboBox5.Items.Add("Current");
                                comboBox5.Items.Add("Temperature 1");
                                comboBox5.Items.Add("Temperature 2");
                                comboBox5.Items.Add("Temperature 3");
                                comboBox5.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox6.Items.Clear();
                                //comboBox6.Text = "";
                                comboBox6.Items.Add("Ending Voltages");
                                comboBox6.Items.Add("Cell 1");
                                comboBox6.Items.Add("Cell 2");
                                comboBox6.Items.Add("Cell 3");
                                comboBox6.Items.Add("Cell 4");
                                comboBox6.Items.Add("Cell 5");
                                comboBox6.Items.Add("Cell 6");
                                comboBox6.Items.Add("Cell 7");
                                comboBox6.Items.Add("Cell 8");
                                comboBox6.Items.Add("Cell 9");
                                comboBox6.Items.Add("Cell 10");
                                comboBox6.Items.Add("Cell 11");
                                comboBox6.Items.Add("Cell 12");
                                comboBox6.Items.Add("Cell 13");
                                comboBox6.Items.Add("Cell 14");
                                comboBox6.Items.Add("Cell 15");
                                comboBox6.Items.Add("Cell 16");
                                comboBox6.Items.Add("Cell 17");
                                comboBox6.Items.Add("Cell 18");
                                comboBox6.Items.Add("Cell 19");
                                break;
                            case "3":
                                // update the cells value
                                if (cell1 == 0) { cell1 = 22; }
                                // Battery combobox
                                comboBox5.Items.Clear();
                                //comboBox5.Text = "";
                                comboBox5.Items.Add("Voltage 1");
                                comboBox5.Items.Add("Voltage 2");
                                comboBox5.Items.Add("Current");
                                comboBox5.Items.Add("Temperature 1");
                                comboBox5.Items.Add("Temperature 2");
                                comboBox5.Items.Add("Temperature 3");
                                comboBox5.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox6.Items.Clear();
                                //comboBox6.Text = "";
                                comboBox6.Items.Add("Ending Voltages");
                                comboBox6.Items.Add("Cell 1");
                                comboBox6.Items.Add("Cell 2");
                                comboBox6.Items.Add("Cell 3");
                                comboBox6.Items.Add("Cell 4");
                                comboBox6.Items.Add("Cell 5");
                                comboBox6.Items.Add("Cell 6");
                                comboBox6.Items.Add("Cell 7");
                                comboBox6.Items.Add("Cell 8");
                                comboBox6.Items.Add("Cell 9");
                                comboBox6.Items.Add("Cell 10");
                                comboBox6.Items.Add("Cell 11");
                                comboBox6.Items.Add("Cell 12");
                                comboBox6.Items.Add("Cell 13");
                                comboBox6.Items.Add("Cell 14");
                                comboBox6.Items.Add("Cell 15");
                                comboBox6.Items.Add("Cell 16");
                                comboBox6.Items.Add("Cell 17");
                                comboBox6.Items.Add("Cell 18");
                                comboBox6.Items.Add("Cell 19");
                                comboBox6.Items.Add("Cell 20");
                                comboBox6.Items.Add("Cell 21");
                                comboBox6.Items.Add("Cell 22");
                                break;
                            case "4":
                                // update the cells value
                                if (cell1 == 0) { cell1 = 21; }
                                // Battery combobox
                                comboBox5.Items.Clear();
                                //comboBox5.Text = "";
                                comboBox5.Items.Add("Voltage 1");
                                comboBox5.Items.Add("Voltage 2");
                                comboBox5.Items.Add("Voltage 3");
                                comboBox5.Items.Add("Current");
                                comboBox5.Items.Add("Temperature 1");
                                comboBox5.Items.Add("Temperature 2");
                                comboBox5.Items.Add("Temperature 3");
                                comboBox5.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox6.Items.Clear();
                                //comboBox6.Text = "";
                                comboBox6.Items.Add("Ending Voltages");
                                comboBox6.Items.Add("Cell 1");
                                comboBox6.Items.Add("Cell 2");
                                comboBox6.Items.Add("Cell 3");
                                comboBox6.Items.Add("Cell 4");
                                comboBox6.Items.Add("Cell 5");
                                comboBox6.Items.Add("Cell 6");
                                comboBox6.Items.Add("Cell 7");
                                comboBox6.Items.Add("Cell 8");
                                comboBox6.Items.Add("Cell 9");
                                comboBox6.Items.Add("Cell 10");
                                comboBox6.Items.Add("Cell 11");
                                comboBox6.Items.Add("Cell 12");
                                comboBox6.Items.Add("Cell 13");
                                comboBox6.Items.Add("Cell 14");
                                comboBox6.Items.Add("Cell 15");
                                comboBox6.Items.Add("Cell 16");
                                comboBox6.Items.Add("Cell 17");
                                comboBox6.Items.Add("Cell 18");
                                comboBox6.Items.Add("Cell 19");
                                comboBox6.Items.Add("Cell 20");
                                comboBox6.Items.Add("Cell 21");
                                break;
                            case "21":
                            case "24":
                                // update the cells value
                                if (cell1 == 0) { cell1 = 21; }
                                // Battery combobox
                                comboBox5.Items.Clear();
                                //comboBox5.Text = "";
                                comboBox5.Items.Add("Voltage");
                                comboBox5.Items.Add("Current");
                                comboBox5.Items.Add("Temperature 1");
                                comboBox5.Items.Add("Temperature 2");
                                comboBox5.Items.Add("Temperature 3");
                                comboBox5.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox6.Items.Clear();
                                //comboBox6.Text = "";
                                comboBox6.Items.Add("Ending Voltages");
                                comboBox6.Items.Add("Cell 1");
                                comboBox6.Items.Add("Cell 2");
                                comboBox6.Items.Add("Cell 3");
                                comboBox6.Items.Add("Cell 4");
                                comboBox6.Items.Add("Cell 5");
                                comboBox6.Items.Add("Cell 6");
                                comboBox6.Items.Add("Cell 7");
                                comboBox6.Items.Add("Cell 8");
                                comboBox6.Items.Add("Cell 9");
                                comboBox6.Items.Add("Cell 10");
                                comboBox6.Items.Add("Cell 11");
                                comboBox6.Items.Add("Cell 12");
                                comboBox6.Items.Add("Cell 13");
                                comboBox6.Items.Add("Cell 14");
                                comboBox6.Items.Add("Cell 15");
                                comboBox6.Items.Add("Cell 16");
                                comboBox6.Items.Add("Cell 17");
                                comboBox6.Items.Add("Cell 18");
                                comboBox6.Items.Add("Cell 19");
                                comboBox6.Items.Add("Cell 20");
                                comboBox6.Items.Add("Cell 21");
                                break;
                            case "22":
                                // update the cells value
                                if (cell1 == 0) { cell1 = 21; }
                                // Battery combobox
                                comboBox5.Items.Clear();
                                //comboBox5.Text = "";
                                comboBox5.Items.Add("Voltage");
                                comboBox5.Items.Add("Current");
                                comboBox5.Items.Add("Temperature 1");
                                comboBox5.Items.Add("Temperature 2");
                                comboBox5.Items.Add("Temperature 3");
                                comboBox5.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox6.Items.Clear();
                                //comboBox6.Text = "";
                                comboBox6.Items.Add("Ending Voltages");
                                comboBox6.Items.Add("Cell 1");
                                comboBox6.Items.Add("Cell 2");
                                comboBox6.Items.Add("Cell 3");
                                comboBox6.Items.Add("Cell 4");
                                comboBox6.Items.Add("Cell 5");
                                comboBox6.Items.Add("Cell 6");
                                comboBox6.Items.Add("Cell 7");
                                comboBox6.Items.Add("Cell 8");
                                comboBox6.Items.Add("Cell 9");
                                comboBox6.Items.Add("Cell 10");
                                comboBox6.Items.Add("Cell 11");
                                comboBox6.Items.Add("Cell 12");
                                comboBox6.Items.Add("Cell 13");
                                comboBox6.Items.Add("Cell 14");
                                comboBox6.Items.Add("Cell 15");
                                comboBox6.Items.Add("Cell 16");
                                comboBox6.Items.Add("Cell 17");
                                comboBox6.Items.Add("Cell 18");
                                comboBox6.Items.Add("Cell 19");
                                comboBox6.Items.Add("Cell 20");
                                comboBox6.Items.Add("Cell 21");
                                comboBox6.Items.Add("Cell 22");
                                break;
                            case "9":
                            case "11":
                                // update the cells value
                                if (cell1 == 0) { cell1 = 20; }
                                // Battery combobox
                                comboBox5.Items.Clear();
                                //comboBox5.Text = "";
                                comboBox5.Items.Add("Voltage 1");
                                comboBox5.Items.Add("Voltage 2");
                                comboBox5.Items.Add("Current");
                                comboBox5.Items.Add("Temperature 1");
                                comboBox5.Items.Add("Temperature 2");
                                comboBox5.Items.Add("Temperature 3");
                                comboBox5.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox6.Items.Clear();
                                //comboBox6.Text = "";
                                comboBox6.Items.Add(" ");
                                break;
                            case "10":
                                // update the cells value
                                if (cell1 == 0) { cell1 = 20; }
                                // Battery combobox
                                comboBox5.Items.Clear();
                                //comboBox5.Text = "";
                                comboBox5.Items.Add("Voltage 1");
                                comboBox5.Items.Add("Voltage 2");
                                comboBox5.Items.Add("Voltage 3");
                                comboBox5.Items.Add("Voltage 4");
                                comboBox5.Items.Add("Current");
                                comboBox5.Items.Add("Temperature 1");
                                comboBox5.Items.Add("Temperature 2");
                                comboBox5.Items.Add("Temperature 3");
                                comboBox5.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox6.Items.Clear();
                                //comboBox6.Text = "";
                                comboBox6.Items.Add(" ");
                                break;
                            default:
                                // update the cells value
                                if (cell1 == 0) { cell1 = 20; }
                                // Battery combobox
                                comboBox5.Items.Clear();
                                //comboBox5.Text = "";
                                comboBox5.Items.Add("Voltage 1");
                                comboBox5.Items.Add("Voltage 2");
                                comboBox5.Items.Add("Voltage 3");
                                comboBox5.Items.Add("Voltage 4");
                                comboBox5.Items.Add("Current 1");
                                comboBox5.Items.Add("Current 2");
                                comboBox5.Items.Add("Temperature 1");
                                comboBox5.Items.Add("Temperature 2");
                                comboBox5.Items.Add("Temperature 3");
                                comboBox5.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox6.Items.Clear();
                                //comboBox6.Text = "";
                                comboBox6.Items.Add("Ending Voltages");
                                comboBox6.Items.Add("Cell 1");
                                comboBox6.Items.Add("Cell 2");
                                comboBox6.Items.Add("Cell 3");
                                comboBox6.Items.Add("Cell 4");
                                comboBox6.Items.Add("Cell 5");
                                comboBox6.Items.Add("Cell 6");
                                comboBox6.Items.Add("Cell 7");
                                comboBox6.Items.Add("Cell 8");
                                comboBox6.Items.Add("Cell 9");
                                comboBox6.Items.Add("Cell 10");
                                comboBox6.Items.Add("Cell 11");
                                comboBox6.Items.Add("Cell 12");
                                comboBox6.Items.Add("Cell 13");
                                comboBox6.Items.Add("Cell 14");
                                comboBox6.Items.Add("Cell 15");
                                comboBox6.Items.Add("Cell 16");
                                comboBox6.Items.Add("Cell 17");
                                comboBox6.Items.Add("Cell 18");
                                comboBox6.Items.Add("Cell 19");
                                comboBox6.Items.Add("Cell 20");
                                comboBox6.Items.Add("Cell 21");
                                comboBox6.Items.Add("Cell 22");
                                comboBox6.Items.Add("Cell 23");
                                comboBox6.Items.Add("Cell 24");
                                break;
                        }// end switch

                        

                    });// end invoke

                    // Now we need to look up some additional info in the Battery table for the graphs....
                    // first get the battry Model...
                    strAccessSelect = @"SELECT BatteryModel FROM WorkOrders WHERE WorkOrderNumber='" + combo1Text + @"'";
                    DataSet wo = new DataSet();

                    myAccessConn = null;
                    // try to open the DB
                    try
                    {
                        myAccessConn = new OleDbConnection(strAccessConn);
                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        return;
                    }
                    // try to open the DB
                    //  now try to access it
                    try
                    {
                        OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                        lock (Main_Form.dataBaseLock)
                        {
                            myAccessConn.Open();
                            myDataAdapter.Fill(wo, "WO");
                            myAccessConn.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        return;
                    }

                    batMod1 = wo.Tables[0].Rows[0][0].ToString();


                    //Now get the nominal voltage and the cell voltage level...
                    strAccessSelect = @"SELECT VOLT,CCVMAX,BTECH,NCELLS FROM BatteriesCustom WHERE BatteryModel='" + batMod1 + @"'";
                    batInfo1 = new DataSet();

                    // try to open the DB
                    //  now try to access it
                    try
                    {
                        OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                        lock (Main_Form.dataBaseLock)
                        {
                            myAccessConn.Open();
                            myDataAdapter.Fill(batInfo1, "Bat");
                            myAccessConn.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        return;
                    }

                    try
                    {
                        NomV1 = (float) GetDouble(batInfo1.Tables[0].Rows[0][0].ToString());
                    }
                    catch
                    {
                        NomV1 = 24;
                    }
                    try
                    {
                        CellV1 = (float) GetDouble(batInfo1.Tables[0].Rows[0][1].ToString());
                    }
                    catch
                    {
                        CellV1 = (float)1.75;
                    }
                    try
                    {
                        technology1 = batInfo1.Tables[0].Rows[0][2].ToString();
                    }
                    catch
                    {
                        technology1 = "NiCd";
                    }
                    try
                    {
                        if (batInfo1.Tables[0].Rows[0][3].ToString() != "")
                        {
                            cell1 = int.Parse(batInfo1.Tables[0].Rows[0][3].ToString());
                        }
                    }
                    catch
                    {
                        //leave as is...
                    }

                    this.Invoke((MethodInvoker)delegate()
                    {

                        comboBox5.SelectedIndex = 0;
                        comboBox6.SelectedIndex = 0;

                        comboBox5.Enabled = true;
                        comboBox6.Enabled = true;
                        if (startup)
                        {
                            startup = false;
                            comboBox5.SelectedIndex = 0;
                        }

                    });// end invoke



                });// end helper thread
            }
        }

        private void comboBox4_SelectedValueChanged(object sender, EventArgs e)
        {

            chart2.Series.Clear();
            tt2.Text = "";

            string combo3Text = comboBox3.Text;
            string combo4Text = comboBox4.Text;
            if (comboBox4.SelectedIndex <= 0) { return; }
            else
            {
                comboBox7.Enabled = false;
                comboBox7.Text = "Loading";
                comboBox8.Enabled = false;
                comboBox8.Text = "Loading";

                // do it on a helper thread!
                ThreadPool.QueueUserWorkItem(s =>
                {
                    // FIRST CLEAR THE OLD DATA SET!
                    graph2Set.Clear();
                    // Open database containing all the battery data....
                    string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                    string strAccessSelect = @"SELECT * FROM ScanData WHERE BWO='" + combo3Text + @"' AND STEP='" + combo4Text.Substring(0, 2) + @"' ORDER BY RDG ASC";

                    //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
                    OleDbConnection myAccessConn = null;
                    // try to open the DB
                    try
                    {
                        myAccessConn = new OleDbConnection(strAccessConn);
                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
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
                            myDataAdapter.Fill(graph2Set, "ScanData");
                            myAccessConn.Close();
                        }


                    }
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        return;
                    }

                    // we have the data table at this point, but we still need to update the graph combo boxes
                    // first we look up the cable used for the test

                    // Open database containing all the battery data....
                    strAccessSelect = @"SELECT CellCableID FROM Tests WHERE WorkOrderNumber='" + combo3Text + @"' AND StepNumber='" + combo4Text.Substring(0, 2) + @"'";


                    DataSet lookForCable = new DataSet();
                    myAccessConn = null;
                    // try to open the DB
                    try
                    {
                        myAccessConn = new OleDbConnection(strAccessConn);
                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
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
                            myDataAdapter.Fill(lookForCable, "Tests");
                            myAccessConn.Close();
                        }
 

                    }
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        });
                        return;
                    }

                    string cellCable = lookForCable.Tables["Tests"].Rows[0][0].ToString();

                    this.Invoke((MethodInvoker)delegate()
                    {
                        // don't let the user look at cell values for SLA bats...
                        if (cellCable != "10")
                        {
                            radioButton4.Enabled = true;
                        }
                        else
                        {
                            radioButton4.Enabled = false;
                            radioButton3.Checked = true;
                        }

                        switch (cellCable)
                        {
                            case "1":
                            case "23":
                                // update the cells value
                                if (cell2 == 0) { cell2 = 20; }
                                // Battery combobox
                                comboBox7.Items.Clear();
                                //comboBox7.Text = "";
                                comboBox7.Items.Add("Voltage");
                                comboBox7.Items.Add("Current");
                                comboBox7.Items.Add("Temperature 1");
                                comboBox7.Items.Add("Temperature 2");
                                comboBox7.Items.Add("Temperature 3");
                                comboBox7.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox8.Items.Clear();
                                //comboBox8.Text = "";
                                comboBox8.Items.Add("Ending Voltages");
                                comboBox8.Items.Add("Cell 1");
                                comboBox8.Items.Add("Cell 2");
                                comboBox8.Items.Add("Cell 3");
                                comboBox8.Items.Add("Cell 4");
                                comboBox8.Items.Add("Cell 5");
                                comboBox8.Items.Add("Cell 6");
                                comboBox8.Items.Add("Cell 7");
                                comboBox8.Items.Add("Cell 8");
                                comboBox8.Items.Add("Cell 9");
                                comboBox8.Items.Add("Cell 10");
                                comboBox8.Items.Add("Cell 11");
                                comboBox8.Items.Add("Cell 12");
                                comboBox8.Items.Add("Cell 13");
                                comboBox8.Items.Add("Cell 14");
                                comboBox8.Items.Add("Cell 15");
                                comboBox8.Items.Add("Cell 16");
                                comboBox8.Items.Add("Cell 17");
                                comboBox8.Items.Add("Cell 18");
                                comboBox8.Items.Add("Cell 19");
                                comboBox8.Items.Add("Cell 20");
                                break;
                            case "2":
                                // update the cells value
                                if (cell2 == 0) { cell2 = 20; }
                                // Battery combobox
                                comboBox7.Items.Clear();
                                //comboBox7.Text = "";
                                comboBox7.Items.Add("Voltage");
                                comboBox7.Items.Add("Current");
                                comboBox7.Items.Add("Temperature 1");
                                comboBox7.Items.Add("Temperature 2");
                                comboBox7.Items.Add("Temperature 3");
                                comboBox7.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox8.Items.Clear();
                                //comboBox8.Text = "";
                                comboBox8.Items.Add("Ending Voltages");
                                comboBox8.Items.Add("Cell 1");
                                comboBox8.Items.Add("Cell 2");
                                comboBox8.Items.Add("Cell 3");
                                comboBox8.Items.Add("Cell 4");
                                comboBox8.Items.Add("Cell 5");
                                comboBox8.Items.Add("Cell 6");
                                comboBox8.Items.Add("Cell 7");
                                comboBox8.Items.Add("Cell 8");
                                comboBox8.Items.Add("Cell 9");
                                comboBox8.Items.Add("Cell 10");
                                comboBox8.Items.Add("Cell 11");
                                comboBox8.Items.Add("Cell 12");
                                comboBox8.Items.Add("Cell 13");
                                comboBox8.Items.Add("Cell 14");
                                comboBox8.Items.Add("Cell 15");
                                comboBox8.Items.Add("Cell 16");
                                comboBox8.Items.Add("Cell 17");
                                comboBox8.Items.Add("Cell 18");
                                comboBox8.Items.Add("Cell 19");
                                break;
                            case "3":
                                // update the cells value
                                if (cell2 == 0) { cell2 = 22; }
                                // Battery combobox
                                comboBox7.Items.Clear();
                                //comboBox7.Text = "";
                                comboBox7.Items.Add("Voltage 1");
                                comboBox7.Items.Add("Voltage 2");
                                comboBox7.Items.Add("Current");
                                comboBox7.Items.Add("Temperature 1");
                                comboBox7.Items.Add("Temperature 2");
                                comboBox7.Items.Add("Temperature 3");
                                comboBox7.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox8.Items.Clear();
                                //comboBox8.Text = "";
                                comboBox8.Items.Add("Ending Voltages");
                                comboBox8.Items.Add("Cell 1");
                                comboBox8.Items.Add("Cell 2");
                                comboBox8.Items.Add("Cell 3");
                                comboBox8.Items.Add("Cell 4");
                                comboBox8.Items.Add("Cell 5");
                                comboBox8.Items.Add("Cell 6");
                                comboBox8.Items.Add("Cell 7");
                                comboBox8.Items.Add("Cell 8");
                                comboBox8.Items.Add("Cell 9");
                                comboBox8.Items.Add("Cell 10");
                                comboBox8.Items.Add("Cell 11");
                                comboBox8.Items.Add("Cell 12");
                                comboBox8.Items.Add("Cell 13");
                                comboBox8.Items.Add("Cell 14");
                                comboBox8.Items.Add("Cell 15");
                                comboBox8.Items.Add("Cell 16");
                                comboBox8.Items.Add("Cell 17");
                                comboBox8.Items.Add("Cell 18");
                                comboBox8.Items.Add("Cell 19");
                                comboBox8.Items.Add("Cell 20");
                                comboBox8.Items.Add("Cell 21");
                                comboBox8.Items.Add("Cell 22");
                                break;
                            case "4":
                                // update the cells value
                                if (cell2 == 0) { cell2 = 21; }
                                // Battery combobox
                                comboBox7.Items.Clear();
                                //comboBox7.Text = "";
                                comboBox7.Items.Add("Voltage 1");
                                comboBox7.Items.Add("Voltage 2");
                                comboBox7.Items.Add("Voltage 3");
                                comboBox7.Items.Add("Current");
                                comboBox7.Items.Add("Temperature 1");
                                comboBox7.Items.Add("Temperature 2");
                                comboBox7.Items.Add("Temperature 3");
                                comboBox7.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox8.Items.Clear();
                                //comboBox8.Text = "";
                                comboBox8.Items.Add("Ending Voltages");
                                comboBox8.Items.Add("Cell 1");
                                comboBox8.Items.Add("Cell 2");
                                comboBox8.Items.Add("Cell 3");
                                comboBox8.Items.Add("Cell 4");
                                comboBox8.Items.Add("Cell 5");
                                comboBox8.Items.Add("Cell 6");
                                comboBox8.Items.Add("Cell 7");
                                comboBox8.Items.Add("Cell 8");
                                comboBox8.Items.Add("Cell 9");
                                comboBox8.Items.Add("Cell 10");
                                comboBox8.Items.Add("Cell 11");
                                comboBox8.Items.Add("Cell 12");
                                comboBox8.Items.Add("Cell 13");
                                comboBox8.Items.Add("Cell 14");
                                comboBox8.Items.Add("Cell 15");
                                comboBox8.Items.Add("Cell 16");
                                comboBox8.Items.Add("Cell 17");
                                comboBox8.Items.Add("Cell 18");
                                comboBox8.Items.Add("Cell 19");
                                comboBox8.Items.Add("Cell 20");
                                comboBox8.Items.Add("Cell 21");
                                break;
                            case "21":
                            case "24":
                                // update the cells value
                                if (cell2 == 0) { cell2 = 21; }
                                // Battery combobox
                                comboBox7.Items.Clear();
                                //comboBox7.Text = "";
                                comboBox7.Items.Add("Voltage");
                                comboBox7.Items.Add("Current");
                                comboBox7.Items.Add("Temperature 1");
                                comboBox7.Items.Add("Temperature 2");
                                comboBox7.Items.Add("Temperature 3");
                                comboBox7.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox8.Items.Clear();
                                //comboBox8.Text = "";
                                comboBox8.Items.Add("Ending Voltages");
                                comboBox8.Items.Add("Cell 1");
                                comboBox8.Items.Add("Cell 2");
                                comboBox8.Items.Add("Cell 3");
                                comboBox8.Items.Add("Cell 4");
                                comboBox8.Items.Add("Cell 5");
                                comboBox8.Items.Add("Cell 6");
                                comboBox8.Items.Add("Cell 7");
                                comboBox8.Items.Add("Cell 8");
                                comboBox8.Items.Add("Cell 9");
                                comboBox8.Items.Add("Cell 10");
                                comboBox8.Items.Add("Cell 11");
                                comboBox8.Items.Add("Cell 12");
                                comboBox8.Items.Add("Cell 13");
                                comboBox8.Items.Add("Cell 14");
                                comboBox8.Items.Add("Cell 15");
                                comboBox8.Items.Add("Cell 16");
                                comboBox8.Items.Add("Cell 17");
                                comboBox8.Items.Add("Cell 18");
                                comboBox8.Items.Add("Cell 19");
                                comboBox8.Items.Add("Cell 20");
                                comboBox8.Items.Add("Cell 21");
                                break;
                            case "22":
                                // update the cells value
                                if (cell2 == 0) { cell2 = 21; }
                                // Battery combobox
                                comboBox7.Items.Clear();
                                //comboBox7.Text = "";
                                comboBox7.Items.Add("Voltage");
                                comboBox7.Items.Add("Current");
                                comboBox7.Items.Add("Temperature 1");
                                comboBox7.Items.Add("Temperature 2");
                                comboBox7.Items.Add("Temperature 3");
                                comboBox7.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox8.Items.Clear();
                                //comboBox8.Text = "";
                                comboBox8.Items.Add("Ending Voltages");
                                comboBox8.Items.Add("Cell 1");
                                comboBox8.Items.Add("Cell 2");
                                comboBox8.Items.Add("Cell 3");
                                comboBox8.Items.Add("Cell 4");
                                comboBox8.Items.Add("Cell 5");
                                comboBox8.Items.Add("Cell 6");
                                comboBox8.Items.Add("Cell 7");
                                comboBox8.Items.Add("Cell 8");
                                comboBox8.Items.Add("Cell 9");
                                comboBox8.Items.Add("Cell 10");
                                comboBox8.Items.Add("Cell 11");
                                comboBox8.Items.Add("Cell 12");
                                comboBox8.Items.Add("Cell 13");
                                comboBox8.Items.Add("Cell 14");
                                comboBox8.Items.Add("Cell 15");
                                comboBox8.Items.Add("Cell 16");
                                comboBox8.Items.Add("Cell 17");
                                comboBox8.Items.Add("Cell 18");
                                comboBox8.Items.Add("Cell 19");
                                comboBox8.Items.Add("Cell 20");
                                comboBox8.Items.Add("Cell 21");
                                comboBox8.Items.Add("Cell 22");
                                break;
                            case "9":
                            case "11":
                                // update the cells value
                                if (cell2 == 0) { cell2 = 20; }
                                // Battery combobox
                                comboBox7.Items.Clear();
                                //comboBox7.Text = "";
                                comboBox7.Items.Add("Voltage 1");
                                comboBox7.Items.Add("Voltage 2");
                                comboBox7.Items.Add("Current");
                                comboBox7.Items.Add("Temperature 1");
                                comboBox7.Items.Add("Temperature 2");
                                comboBox7.Items.Add("Temperature 3");
                                comboBox7.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox8.Items.Clear();
                                //comboBox8.Text = "";
                                comboBox8.Items.Add(" ");
                                break;
                            case "10":
                                // update the cells value
                                if (cell2 == 0) { cell2 = 20; }
                                // Battery combobox
                                comboBox7.Items.Clear();
                                //comboBox7.Text = "";
                                comboBox7.Items.Add("Voltage 1");
                                comboBox7.Items.Add("Voltage 2");
                                comboBox7.Items.Add("Voltage 3");
                                comboBox7.Items.Add("Voltage 4");
                                comboBox7.Items.Add("Current");
                                comboBox7.Items.Add("Temperature 1");
                                comboBox7.Items.Add("Temperature 2");
                                comboBox7.Items.Add("Temperature 3");
                                comboBox7.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox8.Items.Clear();
                                //comboBox8.Text = "";
                                comboBox8.Items.Add(" ");
                                break;
                            default:
                                // update the cells value
                                if (cell2 == 0) { cell2 = 20; }
                                // Battery combobox
                                comboBox7.Items.Clear();
                                //comboBox7.Text = "";
                                comboBox7.Items.Add("Voltage 1");
                                comboBox7.Items.Add("Voltage 2");
                                comboBox7.Items.Add("Voltage 3");
                                comboBox7.Items.Add("Voltage 4");
                                comboBox7.Items.Add("Current 1");
                                comboBox7.Items.Add("Current 2");
                                comboBox7.Items.Add("Temperature 1");
                                comboBox7.Items.Add("Temperature 2");
                                comboBox7.Items.Add("Temperature 3");
                                comboBox7.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox8.Items.Clear();
                                //comboBox8.Text = "";
                                comboBox8.Items.Add("Ending Voltages");
                                comboBox8.Items.Add("Cell 1");
                                comboBox8.Items.Add("Cell 2");
                                comboBox8.Items.Add("Cell 3");
                                comboBox8.Items.Add("Cell 4");
                                comboBox8.Items.Add("Cell 5");
                                comboBox8.Items.Add("Cell 6");
                                comboBox8.Items.Add("Cell 7");
                                comboBox8.Items.Add("Cell 8");
                                comboBox8.Items.Add("Cell 9");
                                comboBox8.Items.Add("Cell 10");
                                comboBox8.Items.Add("Cell 11");
                                comboBox8.Items.Add("Cell 12");
                                comboBox8.Items.Add("Cell 13");
                                comboBox8.Items.Add("Cell 14");
                                comboBox8.Items.Add("Cell 15");
                                comboBox8.Items.Add("Cell 16");
                                comboBox8.Items.Add("Cell 17");
                                comboBox8.Items.Add("Cell 18");
                                comboBox8.Items.Add("Cell 19");
                                comboBox8.Items.Add("Cell 20");
                                comboBox8.Items.Add("Cell 21");
                                comboBox8.Items.Add("Cell 22");
                                comboBox8.Items.Add("Cell 23");
                                comboBox8.Items.Add("Cell 24");
                                break;
                        }// end switch

                        

                        // Now we need to look up some additional info in the Battery table for the graphs....
                        // first get the battry Model...
                        strAccessSelect = @"SELECT BatteryModel FROM WorkOrders WHERE WorkOrderNumber='" + combo3Text + @"'";
                        DataSet wo = new DataSet();
                        myAccessConn = null;
                        // try to open the DB
                        try
                        {
                            myAccessConn = new OleDbConnection(strAccessConn);
                        }
                        catch (Exception ex)
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            return;
                        }
                        // try to open the DB
                        //  now try to access it
                        try
                        {
                            OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                            lock (Main_Form.dataBaseLock)
                            {
                                myAccessConn.Open();
                                myDataAdapter.Fill(wo, "WO");
                                myAccessConn.Close();
                            }

                        }
                        catch (Exception ex)
                        {
                            myAccessConn.Close();
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            return;
                        }

                        batMod2 = wo.Tables[0].Rows[0][0].ToString();


                        //Now get the nominal voltage and the cell voltage level...
                        strAccessSelect = @"SELECT VOLT,CCVMAX,BTECH,NCELLS FROM BatteriesCustom WHERE BatteryModel='" + batMod2 + @"'";
                        batInfo2 = new DataSet();
                        myAccessConn = null;
                        // try to open the DB
                        try
                        {
                            myAccessConn = new OleDbConnection(strAccessConn);
                        }
                        catch (Exception ex)
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            return;
                        }
                        // try to open the DB
                        //  now try to access it
                        try
                        {
                            OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                            lock (Main_Form.dataBaseLock)
                            {
                                myAccessConn.Open();
                                myDataAdapter.Fill(batInfo2, "Bat");
                                myAccessConn.Close();
                            }

                        }
                        catch (Exception ex)
                        {
                            myAccessConn.Close();
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            });
                            return;
                        }

                        try
                        {
                            NomV2 = (float) GetDouble(batInfo2.Tables[0].Rows[0][0].ToString());
                        }
                        catch 
                        {
                            NomV2 = 24;
                        }
                        try
                        {
                            CellV2 = (float) GetDouble(batInfo2.Tables[0].Rows[0][1].ToString());
                        }
                        catch
                        {
                            CellV2 = (float) 1.75;
                        }
                        try
                        {
                            technology2 = batInfo2.Tables[0].Rows[0][2].ToString();
                        }
                        catch
                        {
                            technology2 = "NiCd";
                        }
                        try
                        {
                            if (batInfo2.Tables[0].Rows[0][3].ToString() != "")
                            {
                                cell2 = int.Parse(batInfo2.Tables[0].Rows[0][3].ToString());
                            }
                        }
                        catch
                        {
                            //leave as is...
                        }


                        comboBox7.Enabled = true;
                        comboBox8.Enabled = true;

                        comboBox7.SelectedIndex = 0;
                        comboBox8.SelectedIndex = 0;


                        if (startup2)
                        {
                            startup2 = false;
                            comboBox7.SelectedIndex = 0;
                        }
                    });// end invoke
                });// end helper thread
            }
        }

        private void comboBox5_SelectedValueChanged(object sender, EventArgs e)
        {

            type1 = comboBox2.Text;

            try
            {
                int q;
                // only do something if the radio button is selected
                if (radioButton1.Checked == false || comboBox5.SelectedIndex < 0) { return; }
                // Here we will look at the Value selected and then plot graph1Set

                //find out which graph to plot from the selected text
                switch (comboBox5.Text)
                {
                    case "Voltage":
                    case "Voltage 1":
                        q = 10;
                        break;
                    case "Voltage 2":
                        q = 11;
                        break;
                    case "Voltage 3":
                        q = 12;
                        break;
                    case "Voltage 4":
                        q = 13;
                        break;
                    case "Current":
                    case "Current 1":
                        q = 8;
                        break;
                    case "Current 2":
                        q = 9;
                        break;
                    case "Temperature 1":
                        q = 38;
                        break;
                    case "Temperature 2":
                        q = 39;
                        break;
                    case "Temperature 3":
                        q = 40;
                        break;
                    case "Temperature 4":
                        q = 41;
                        break;
                    default:
                        q = 7;
                        break;
                }

                //we need to graph the col 7 as time and q as the value
                this.chart1.Series.Clear();
                var series1 = new System.Windows.Forms.DataVisualization.Charting.Series
                {
                    Name = "Series1",
                    Color = System.Drawing.Color.Green,
                    IsVisibleInLegend = false,
                    IsXValueIndexed = true,
                    ChartType = SeriesChartType.Column,
                    BorderColor = System.Drawing.Color.DarkGray,
                    BorderWidth = 1
                };
                this.chart1.Series.Add(series1);



                // pad with zero Vals to help with the look of the plot...
                // first get the interval and total points
                double interval = 1;
                int points = 1;

                switch (comboBox2.Text.Substring(5))
                {
                    case "As Received":
                        interval = 1 / 30;
                        points = 3;
                        chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Full Charge-6":
                        interval = 5;
                        points = 73;
                        chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Full Charge-4":
                        interval = 4;
                        points = 61;
                        chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Top Charge-4":
                        interval = 4;
                        points = 61;
                        chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Top Charge-2":
                        interval = 3;
                        points = 41;
                        chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Top Charge-1":
                        interval = 1;
                        points = 61;
                        chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Constant Voltage":
                        interval = 5;
                        points = 73;
                        chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Capacity-1":
                    case "Capacity":
                        interval = 1;
                        points = 61;
                        chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Discharge":
                        interval = 1;
                        points = 61;
                        chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Slow Charge-14":
                        interval = 12;
                        points = 73;
                        chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "SlowCharge-16":
                    case "Shorting-16":
                        interval = 16;
                        points = 61;
                        chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    default:
                        //custom cap and charge get the default...
                        //Look it up!
                        string tempString = comboBox2.Text.Substring(5);
                        if (tempString.Contains("Combo:") && tempString.Contains("("))
                        {
                            tempString = tempString.Substring(tempString.IndexOf("("), tempString.IndexOf(")") - tempString.IndexOf("("));
                            tempString = tempString.Substring(tempString.IndexOf(" ") + 1);

                        }

                        for (int i = 0; i < customTestParams.Rows.Count; i++)
                        {
                            if (customTestParams.Rows[i][1].ToString() == tempString)
                            {
                                interval = ((int)customTestParams.Rows[i][4]) / 60.0;
                                points = (int)customTestParams.Rows[i][3];
                                chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0.0";
                                break;
                            }
                        }

                        chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0.0";
                        break;
                }
                
                // now add points
                for (int i = 0; i < graph1Set.Tables[0].Rows.Count; i++)
                {
                    series1.Points.AddXY(Math.Round(GetDouble(graph1Set.Tables[0].Rows[i][7].ToString()) * 1440, ((interval < 1) ? 1 : 0)), graph1Set.Tables[0].Rows[i][q]);
                    // color test
                    series1.Points[i].Color = pointColor(technology1, cell1, GetDouble(graph1Set.Tables[0].Rows[i][q].ToString()), type1);
                }


                if (graph1Set.Tables[0].Rows.Count <= points - 1)
                {
                    for (int i = graph1Set.Tables[0].Rows.Count; i <= points - 1; i++)
                    {
                        series1.Points.AddXY(i * interval, 0);
                    }
                }

                tt1.Text = "Work Order:  " + comboBox1.Text + "     Test:  " + comboBox2.Text + "     Date:  " + graph1Set.Tables[0].Rows[0][4].ToString();
                chart1.Invalidate();
                chart1.ChartAreas[0].RecalculateAxesScale();
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate()
                {
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        }
        private void comboBox6_SelectedValueChanged(object sender, EventArgs e)
        {
            type1 = comboBox2.Text;
            try
            {
                int q;
                // only do something if the radio button is selected
                if (radioButton2.Checked == false || comboBox6.SelectedIndex < 0) { return; }
                // Here we will look at the Value selected and then plot graph1Set

                //find out which graph to plot from the selected text
                int numCells = cell1;

                switch (comboBox6.Text)
                {
                    case "Ending Voltages":
                        q = 999;
                        break;
                    case "Cell 1":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells; }
                        else { q = 14; }
                        break;
                    case "Cell 2":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 1; }
                        else { q = 15; }
                        break;
                    case "Cell 3":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 2; }
                        else { q = 16; }
                        break;
                    case "Cell 4":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 3; }
                        else { q = 17; }
                        break;
                    case "Cell 5":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 4; }
                        else { q = 18; }
                        break;
                    case "Cell 6":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 5; }
                        else { q = 19; }
                        break;
                    case "Cell 7":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 6; }
                        else { q = 20; }
                        break;
                    case "Cell 8":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 7; }
                        else { q = 21; }
                        break;
                    case "Cell 9":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 8; }
                        else { q = 22; }
                        break;
                    case "Cell 10":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 9; }
                        else { q = 23; }
                        break;
                    case "Cell 11":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 10; }
                        else { q = 24; }
                        break;
                    case "Cell 12":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 11; }
                        else { q = 25; }
                        break;
                    case "Cell 13":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 12; }
                        else { q = 26; }
                        break;
                    case "Cell 14":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 13; }
                        else { q = 27; }
                        break;
                    case "Cell 15":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 14; }
                        else { q = 28; }
                        break;
                    case "Cell 16":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 15; }
                        else { q = 29; }
                        break;
                    case "Cell 17":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 16; }
                        else { q = 30; }
                        break;
                    case "Cell 18":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 17; }
                        else { q = 31; }
                        break;
                    case "Cell 19":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 18; }
                        else { q = 32; }
                        break;
                    case "Cell 20":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 19; }
                        else { q = 33; }
                        break;
                    case "Cell 21":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 20; }
                        else { q = 34; }
                        break;
                    case "Cell 22":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 21; }
                        else { q = 35; }
                        break;
                    case "Cell 23":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 22; }
                        else { q = 36; }
                        break;
                    case "Cell 24":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 23; }
                        else { q = 37; }
                        break;
                    default:
                        q = 999;
                        break;
                }

                //we need to graph the col 7 as time and q as the value
                this.chart1.Series.Clear();
                var series1 = new System.Windows.Forms.DataVisualization.Charting.Series
                {
                    Name = "Series1",
                    Color = System.Drawing.Color.Green,
                    IsVisibleInLegend = false,
                    IsXValueIndexed = true,
                    ChartType = SeriesChartType.Column,
                    BorderColor = System.Drawing.Color.DarkGray,
                    BorderWidth = 1
                };
                this.chart1.Series.Add(series1);

                if (q == 999)
                {
                    for (int i = 0; i < cell1; i++)
                    {
                        if (GlobalVars.Pos2Neg == false)
                        {
                            series1.Points.AddXY(i + 1, graph1Set.Tables[0].Rows[graph1Set.Tables[0].Rows.Count - 1][i + 14]);
                            // color test
                            series1.Points[i].Color = pointColor(technology1, 1, GetDouble(graph1Set.Tables[0].Rows[graph1Set.Tables[0].Rows.Count - 1][i + 14].ToString()), type1);
                        }
                        else
                        {
                            series1.Points.AddXY(i + 1, graph1Set.Tables[0].Rows[graph1Set.Tables[0].Rows.Count - 1][cell1 - i - 1 + 14]);
                            // color test
                            series1.Points[i].Color = pointColor(technology1, 1, GetDouble(graph1Set.Tables[0].Rows[graph1Set.Tables[0].Rows.Count - 1][cell1 - i - 1 + 14].ToString()), type1);
                        }
                    }
                }
                else
                {




                    // pad with zero Vals to help with the look of the plot...
                    // first get the interval and total points
                    double interval = 1;
                    int points = 1;

                    switch (comboBox2.Text.Substring(5))
                    {
                        case "As Received":
                            interval = 1 / 30;
                            points = 3;
                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Full Charge-6":
                            interval = 5;
                            points = 73;
                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Full Charge-4":
                            interval = 4;
                            points = 61;
                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Top Charge-4":
                            interval = 4;
                            points = 61;
                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Top Charge-2":
                            interval = 3;
                            points = 41;
                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Top Charge-1":
                            interval = 1;
                            points = 61;
                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Constant Voltage":
                            interval = 5;
                            points = 73;
                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Capacity-1":
                        case "Capacity":
                            interval = 1;
                            points = 61;
                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Discharge":
                            interval = 1;
                            points = 61;
                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Slow Charge-14":
                            interval = 12;
                            points = 73;
                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "SlowCharge-16":
                        case "Shorting-16":
                            interval = 16;
                            points = 61;
                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        default:
                            //custom cap and charge get the default...
                            //Look it up!
                            string tempString = comboBox2.Text.Substring(5);
                            if (tempString.Contains("Combo:") && tempString.Contains("("))
                            {
                                tempString = tempString.Substring(tempString.IndexOf("("), tempString.IndexOf(")") - tempString.IndexOf("("));
                                tempString = tempString.Substring(tempString.IndexOf(" ") + 1);

                            }

                            for (int i = 0; i < customTestParams.Rows.Count; i++)
                            {
                                if (customTestParams.Rows[i][1].ToString() == tempString)
                                {
                                    interval = ((int)customTestParams.Rows[i][4]) / 60.0;
                                    points = (int)customTestParams.Rows[i][3];
                                    chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0.0";
                                    break;
                                }
                            }

                            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "0.0";
                            break;
                    }

                    for (int i = 0; i < graph1Set.Tables[0].Rows.Count; i++)
                    {
                        series1.Points.AddXY(Math.Round(GetDouble(graph1Set.Tables[0].Rows[i][7].ToString()) * 1440, ((interval < 1) ? 1 : 0)), graph1Set.Tables[0].Rows[i][q]);
                        // color test
                        series1.Points[i].Color = pointColor(technology1, 1, GetDouble(graph1Set.Tables[0].Rows[i][q].ToString()), type1);
                    }

                    if (graph1Set.Tables[0].Rows.Count <= points - 1)
                    {
                        for (int i = graph1Set.Tables[0].Rows.Count; i <= points - 1; i++)
                        {
                            series1.Points.AddXY(i * interval, 0);
                        }
                    }
                }

                tt1.Text = "Work Order:  " + comboBox1.Text + "     Test:  " + comboBox2.Text + "     Date:  " + graph1Set.Tables[0].Rows[0][4].ToString();
                chart1.Invalidate();
                chart1.ChartAreas[0].RecalculateAxesScale();


            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate()
                {
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }

        }

        /// <summary>
        /// This function returns color to use in the plots produced by the program
        /// </summary>
        /// <param name="tech">this is the type of cell being looked at (NiCd, Lead Acid, ..)</param>
        /// <param name="Cells">this is the number of cells in the battery</param>
        /// <param name="Value">the reading value to assign a color to</param>
        /// <param name="type">Indicates if the plot is a Charge, Discharge, Capacity, or As Received plot</param>
        /// <returns></returns>
        private Color pointColor(string tech, int Cells, double Value, string type)
        {
            if (radioButton1.Checked == true)
            {
                switch (comboBox5.Text)
                {
                    case "Current":
                    case "Current 1":
                    case "Current 2":
                        return System.Drawing.Color.Blue;
                    case "Temperature 1":
                    case "Temperature 2":
                    case "Temperature 3":
                    case "Temperature 4":
                        return System.Drawing.Color.LightSeaGreen;
                    default:
                        break;
                }
            }

            // test_type is the type of test we are generating the colors for

            // Three types of batteries (NiCd, SLA and NiCd ULM) and two directions (charge discharge)

            // normal vented NiCds
            double Min1 = 0;
            double Min2 = 0;
            double Min3 = 0;
            double Min4 = 0;
            double Max = 0;

            if (CellV1 == 0)
            {
                CellV1 = (float)1.75;
            }

            switch (tech)
            {
                case "NiCd":
                    // Discharge
                    if (type.Contains("As Received") || type.Contains("Cap") || type.Contains("Discharge") || type.Contains("Shorting") || type == "")
                    {
                        Min4 = 1 * Cells;
                        Max = 1.05 * Cells;

                        if (Value > Max) { return Color.Green; }
                        else if (Value > Min4) { return Color.Orange; }
                        else { return Color.Red; }

                    }
                    // Charge
                    else
                    {
                        Min1 = 0.25 * Cells;
                        Min2 = 1.5 * Cells;
                        Min3 = 1.55 * Cells;
                        Max = CellV1 * Cells;

                        if (Value > Max) 
                        { 
                            return Color.Red; 
                        }
                        else if (Value > Min3) { return Color.Green; }
                        else if (Value > Min2) { return Color.Orange; }
                        else if (Value > Min1) { return Color.Yellow; }
                        else { return Color.Red; }
                    }
                case "Sealed Lead Acid":
                    // Discharge
                    if (type.Contains("As Received") || type.Contains("Capacity-1") || type.Contains("Discharge") || type.Contains("Custom Cap") || type.Contains("Shorting") || type == "")
                    {

                        Min4 = (20.0 / 24) * NomV1;
                        Max = (21.0 / 24) * NomV1;

                        if (Value > Max) { return Color.Green; }
                        else if (Value > Min4) { return Color.Orange; }
                        else { return Color.Red; }

                    }
                    // Charge
                    else
                    {
                        // always Blue!
                        return Color.Blue;
                    }
                case "NiCd ULM":
                    // Discharge
                    if (type.Contains("As Received") || type.Contains("Capacity-1") || type.Contains("Discharge") || type.Contains("Custom Cap") || type.Contains("Shorting") || type == "")
                    {
                        Min4 = 1 * Cells;
                        Max = 1.05 * Cells;

                        if (Value > Max) { return Color.Green; }
                        else if (Value > Min4) { return Color.Orange; }
                        else { return Color.Red; }

                    }
                    // Charge
                    else
                    {
                        Min1 = 0.25 * Cells;
                        Min2 = 1.55 * Cells;
                        Min3 = 1.6 * Cells;
                        Max = ((-1 == CellV1) ? 1.82 : CellV1) * Cells;

                        if (Value > Max) { return Color.Red; }
                        else if (Value > Min3) { return Color.Green; }
                        else if (Value > Min2) { return Color.Orange; }
                        else if (Value > Min1) { return Color.Yellow; }
                        else { return Color.Red; }
                    }
            }

            // we'll return a purple if everything goes wrong
            return System.Drawing.Color.Purple;


        }
        private Color pointColor2(string tech, int Cells, double Value, string type)
        {
            if (radioButton3.Checked == true)
            {
                switch (comboBox7.Text)
                {
                    case "Current":
                    case "Current 1":
                    case "Current 2":
                        return System.Drawing.Color.Blue;
                    case "Temperature 1":
                    case "Temperature 2":
                    case "Temperature 3":
                    case "Temperature 4":
                        return System.Drawing.Color.LightSeaGreen;
                    default:
                        break;
                }
            }

            // test_type is the type of test we are generating the colors for

            // Three types of batteries (NiCd, SLA and NiCd ULM) and two directions (charge discharge)

            // normal vented NiCds
            double Min1 = 0;
            double Min2 = 0;
            double Min3 = 0;
            double Min4 = 0;
            double Max = 0;

            if (CellV2 == 0)
            {
                CellV2 = (float)1.75;
            }

            switch (tech)
            {
                case "NiCd":
                    // Discharge
                    if (type.Contains("As Received") || type.Contains("Cap") || type.Contains("Discharge") || type.Contains("Shorting") || type == "")
                    {
                        Min4 = 1 * Cells;
                        Max = 1.05 * Cells;

                        if (Value > Max) { return Color.Green; }
                        else if (Value > Min4) { return Color.Orange; }
                        else { return Color.Red; }

                    }
                    // Charge
                    else
                    {
                        Min1 = 0.25 * Cells;
                        Min2 = 1.5 * Cells;
                        Min3 = 1.55 * Cells;
                        Max = CellV2 * Cells;

                        if (Value > Max) { return Color.Red; }
                        else if (Value > Min3) { return Color.Green; }
                        else if (Value > Min2) { return Color.Orange; }
                        else if (Value > Min1) { return Color.Yellow; }
                        else { return Color.Red; }
                    }
                case "Sealed Lead Acid":
                    // Discharge
                    if (type.Contains("As Received") || type.Contains("Capacity-1") || type.Contains("Discharge") || type.Contains("Custom Cap") || type.Contains("Shorting") || type == "")
                    {

                        Min4 = (20.0 / 24) * NomV2;
                        Max = (21.0 / 24) * NomV2;

                        if (Value > Max) { return Color.Green; }
                        else if (Value > Min4) { return Color.Orange; }
                        else { return Color.Red; }

                    }
                    // Charge
                    else
                    {
                        // always Blue!
                        return Color.Blue;
                    }
                case "NiCd ULM":
                    // Discharge
                    if (type.Contains("As Received") || type.Contains("Capacity-1") || type.Contains("Discharge") || type.Contains("Custom Cap") || type.Contains("Shorting") || type == "")
                    {
                        Min4 = 1.0 * Cells;
                        Max = 1.05 * Cells;

                        if (Value > Max) { return Color.Green; }
                        else if (Value > Min4) { return Color.Orange; }
                        else { return Color.Red; }

                    }
                    // Charge
                    else
                    {
                        Min1 = 0.25 * Cells;
                        Min2 = 1.55 * Cells;
                        Min3 = 1.6 * Cells;
                        Max = CellV2 * Cells;

                        if (Value > Max) { return Color.Red; }
                        else if (Value > Min3) { return Color.Green; }
                        else if (Value > Min2) { return Color.Orange; }
                        else if (Value > Min1) { return Color.Yellow; }
                        else { return Color.Red; }
                    }
            }

            // we'll return a purple if everything goes wrong
            return System.Drawing.Color.Purple;


        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                comboBox5_SelectedValueChanged(radioButton1, null);
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                comboBox6_SelectedValueChanged(radioButton1, null);
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            loadWorkOrderLists();
        }

        private void comboBox7_SelectedValueChanged(object sender, EventArgs e)
        {
            type2 = comboBox4.Text;

            try
            {
                int q;
                // only do something if the radio button is selected
                if (radioButton3.Checked == false || comboBox7.SelectedIndex < 0) { return; }
                // Here we will look at the Value selected and then plot graph1Set

                //find out which graph to plot from the selected text
                switch (comboBox7.Text)
                {
                    case "Voltage":
                    case "Voltage 1":
                        q = 10;
                        break;
                    case "Voltage 2":
                        q = 11;
                        break;
                    case "Voltage 3":
                        q = 12;
                        break;
                    case "Voltage 4":
                        q = 13;
                        break;
                    case "Current":
                    case "Current 1":
                        q = 8;
                        break;
                    case "Current 2":
                        q = 9;
                        break;
                    case "Temperature 1":
                        q = 38;
                        break;
                    case "Temperature 2":
                        q = 39;
                        break;
                    case "Temperature 3":
                        q = 40;
                        break;
                    case "Temperature 4":
                        q = 41;
                        break;
                    default:
                        q = 7;
                        break;
                }

                //we need to graph the col 7 as time and q as the value
                this.chart2.Series.Clear();
                var series2 = new System.Windows.Forms.DataVisualization.Charting.Series
                {
                    Name = "Series2",
                    Color = System.Drawing.Color.Green,
                    IsVisibleInLegend = false,
                    IsXValueIndexed = true,
                    ChartType = SeriesChartType.Column,
                    BorderColor = System.Drawing.Color.DarkGray,
                    BorderWidth = 1
                };
                this.chart2.Series.Add(series2);



                // pad with zero Vals to help with the look of the plot...
                // first get the interval and total points
                double interval = 1;
                int points = 1;

                switch (comboBox4.Text.Substring(5))
                {
                    case "As Received":
                        interval = 1 / 30;
                        points = 3;
                        chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Full Charge-6":
                        interval = 5;
                        points = 73;
                        chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Full Charge-4":
                        interval = 4;
                        points = 61;
                        chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Top Charge-4":
                        interval = 4;
                        points = 61;
                        chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Top Charge-2":
                        interval = 3;
                        points = 41;
                        chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Top Charge-1":
                        interval = 1;
                        points = 61;
                        chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Constant Voltage":
                        interval = 5;
                        points = 73;
                        chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Capacity-1":
                    case "Capacity":
                        interval = 1;
                        points = 61;
                        chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Discharge":
                        interval = 1;
                        points = 61;
                        chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "Slow Charge-14":
                        interval = 12;
                        points = 73;
                        chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    case "SlowCharge-16":
                    case "Shorting-16":
                        interval = 16;
                        points = 61;
                        chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                        break;
                    default:
                        //custom cap and charge get the default...
                        //Look it up!
                        string tempString = comboBox2.Text.Substring(5);
                        if (tempString.Contains("Combo:") && tempString.Contains("("))
                        {
                            tempString = tempString.Substring(tempString.IndexOf("("), tempString.IndexOf(")") - tempString.IndexOf("("));
                            tempString = tempString.Substring(tempString.IndexOf(" ") + 1);

                        }

                        for (int i = 0; i < customTestParams.Rows.Count; i++)
                        {
                            if (customTestParams.Rows[i][1].ToString() == tempString)
                            {
                                interval = ((int)customTestParams.Rows[i][4]) / 60.0;
                                points = (int)customTestParams.Rows[i][3];
                                chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0.0";
                                break;
                            }
                        }

                        chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0.0";
                        break;
                }


                for (int i = 0; i < graph2Set.Tables[0].Rows.Count; i++)
                {
                    series2.Points.AddXY(Math.Round(GetDouble(graph2Set.Tables[0].Rows[i][7].ToString()) * 1440, ((interval < 1) ? 1 : 0)), graph2Set.Tables[0].Rows[i][q]);
                    // color test
                    series2.Points[i].Color = pointColor2(technology2, cell2, GetDouble(graph2Set.Tables[0].Rows[i][q].ToString()), type2);
                }

                if (graph2Set.Tables[0].Rows.Count <= points - 1)
                {
                    for (int i = graph2Set.Tables[0].Rows.Count; i <= points - 1; i++)
                    {
                        series2.Points.AddXY(i * interval, 0);
                    }
                }

                tt2.Text = "Work Order:  " + comboBox3.Text + "     Test:  " + comboBox4.Text + "     Date:  " + graph2Set.Tables[0].Rows[0][4].ToString();
                chart2.Invalidate();
                chart2.ChartAreas[0].RecalculateAxesScale();
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate()
                {
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }

        }

        private void comboBox8_SelectedValueChanged(object sender, EventArgs e)
        {
            type2 = comboBox4.Text;

            try
            {
                int q;
                // only do something if the radio button is selected
                if (radioButton4.Checked == false || comboBox8.SelectedIndex < 0) { return; }
                // Here we will look at the Value selected and then plot graph1Set

                int numCells = cell2;
                //find out which graph to plot from the selected text
                switch (comboBox8.Text)
                {
                    case "Ending Voltages":
                        q = 999;
                        break;
                    case "Cell 1":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells; }
                        else { q = 14; }
                        break;
                    case "Cell 2":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 1; }
                        else { q = 15; }
                        break;
                    case "Cell 3":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 2; }
                        else { q = 16; }
                        break;
                    case "Cell 4":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 3; }
                        else { q = 17; }
                        break;
                    case "Cell 5":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 4; }
                        else { q = 18; }
                        break;
                    case "Cell 6":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 5; }
                        else { q = 19; }
                        break;
                    case "Cell 7":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 6; }
                        else { q = 20; }
                        break;
                    case "Cell 8":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 7; }
                        else { q = 21; }
                        break;
                    case "Cell 9":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 8; }
                        else { q = 22; }
                        break;
                    case "Cell 10":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 9; }
                        else { q = 23; }
                        break;
                    case "Cell 11":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 10; }
                        else { q = 24; }
                        break;
                    case "Cell 12":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 11; }
                        else { q = 25; }
                        break;
                    case "Cell 13":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 12; }
                        else { q = 26; }
                        break;
                    case "Cell 14":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 13; }
                        else { q = 27; }
                        break;
                    case "Cell 15":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 14; }
                        else { q = 28; }
                        break;
                    case "Cell 16":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 15; }
                        else { q = 29; }
                        break;
                    case "Cell 17":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 16; }
                        else { q = 30; }
                        break;
                    case "Cell 18":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 17; }
                        else { q = 31; }
                        break;
                    case "Cell 19":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 18; }
                        else { q = 32; }
                        break;
                    case "Cell 20":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 19; }
                        else { q = 33; }
                        break;
                    case "Cell 21":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 20; }
                        else { q = 34; }
                        break;
                    case "Cell 22":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 21; }
                        else { q = 35; }
                        break;
                    case "Cell 23":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 22; }
                        else { q = 36; }
                        break;
                    case "Cell 24":
                        if (GlobalVars.Pos2Neg) { q = 13 + numCells - 23; }
                        else { q = 37; }
                        break;
                    default:
                        q = 999;
                        break;
                }

                //we need to graph the col 7 as time and q as the value
                this.chart2.Series.Clear();
                var series2 = new System.Windows.Forms.DataVisualization.Charting.Series
                {
                    Name = "Series2",
                    Color = System.Drawing.Color.Green,
                    IsVisibleInLegend = false,
                    IsXValueIndexed = true,
                    ChartType = SeriesChartType.Column,
                    BorderColor = System.Drawing.Color.DarkGray,
                    BorderWidth = 1
                };
                this.chart2.Series.Add(series2);

                if (q == 999)
                {
                    for (int i = 0; i < cell2; i++)
                    {
                        if (GlobalVars.Pos2Neg == false)
                        {
                            series2.Points.AddXY(i + 1, graph2Set.Tables[0].Rows[graph2Set.Tables[0].Rows.Count - 1][i + 14]);
                            // color test
                            series2.Points[i].Color = pointColor2(technology2, 1, GetDouble(graph2Set.Tables[0].Rows[graph2Set.Tables[0].Rows.Count - 1][i + 14].ToString()), type2);
                        }
                        else
                        {
                            series2.Points.AddXY(i + 1, graph2Set.Tables[0].Rows[graph2Set.Tables[0].Rows.Count - 1][cell2 - i - 1 + 14]);
                            // color test
                            series2.Points[i].Color = pointColor2(technology2, 1, GetDouble(graph2Set.Tables[0].Rows[graph2Set.Tables[0].Rows.Count - 1][cell2 - i - 1 + 14].ToString()), type2);
                        }
                    
                    }
                }
                else
                {


                    // pad with zero Vals to help with the look of the plot...
                    // first get the interval and total points
                    double interval = 1;
                    int points = 1;

                    switch (comboBox4.Text.Substring(5))
                    {
                        case "As Received":
                            interval = 1 / 30;
                            points = 3;
                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Full Charge-6":
                            interval = 5;
                            points = 73;
                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Full Charge-4":
                            interval = 4;
                            points = 61;
                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Top Charge-4":
                            interval = 4;
                            points = 61;
                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Top Charge-2":
                            interval = 3;
                            points = 41;
                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Top Charge-1":
                            interval = 1;
                            points = 61;
                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Constant Voltage":
                            interval = 5;
                            points = 73;
                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Capacity-1":
                        case "Capacity":
                            interval = 1;
                            points = 61;
                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Discharge":
                            interval = 1;
                            points = 61;
                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "Slow Charge-14":
                            interval = 12;
                            points = 73;
                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        case "SlowCharge-16":
                        case "Shorting-16":
                            interval = 16;
                            points = 61;
                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0";
                            break;
                        default:
                            //custom cap and charge get the default...
                            //Look it up!
                            string tempString = comboBox2.Text.Substring(5);
                            if (tempString.Contains("Combo:") && tempString.Contains("("))
                            {
                                tempString = tempString.Substring(tempString.IndexOf("("), tempString.IndexOf(")") - tempString.IndexOf("("));
                                tempString = tempString.Substring(tempString.IndexOf(" ") + 1);

                            }

                            for (int i = 0; i < customTestParams.Rows.Count; i++)
                            {
                                if (customTestParams.Rows[i][1].ToString() == tempString)
                                {
                                    interval = ((int)customTestParams.Rows[i][4]) / 60.0;
                                    points = (int)customTestParams.Rows[i][3];
                                    chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0.0";
                                    break;
                                }
                            }

                            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "0.0";
                            break;
                    }

                    for (int i = 0; i < graph2Set.Tables[0].Rows.Count; i++)
                    {
                        series2.Points.AddXY(Math.Round(GetDouble(graph2Set.Tables[0].Rows[i][7].ToString()) * 1440, ((interval < 1) ? 1 : 0)), graph2Set.Tables[0].Rows[i][q]);
                        // color test
                        series2.Points[i].Color = pointColor2(technology2, 1, GetDouble(graph2Set.Tables[0].Rows[i][q].ToString()), type2);
                    }

                    if (graph2Set.Tables[0].Rows.Count <= points - 1)
                    {
                        for (int i = graph2Set.Tables[0].Rows.Count; i <= points - 1; i++)
                        {
                            series2.Points.AddXY(i * interval, 0);
                        }
                    }
                }

                tt2.Text = "Work Order:  " + comboBox3.Text + "     Test:  " + comboBox4.Text + "     Date:  " + graph2Set.Tables[0].Rows[0][4].ToString();
                chart2.Invalidate();
                chart2.ChartAreas[0].RecalculateAxesScale();
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate()
                {
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }


        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked == true)
            {
                comboBox7_SelectedValueChanged(radioButton3, null);
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked == true)
            {
                comboBox8_SelectedValueChanged(radioButton4, null);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PrintDialog MyPrintDialog = new PrintDialog();
            if (MyPrintDialog.ShowDialog() == DialogResult.OK)
            {
                System.Drawing.Printing.PrintDocument doc = new System.Drawing.Printing.PrintDocument();
                doc.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(doc_PrintPage);
                doc.DefaultPageSettings.Landscape = true;
                doc.Print();
            }
        }

        private void doc_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Bitmap bmp = new Bitmap(panel1.Width, panel1.Height, panel1.CreateGraphics());
            panel1.DrawToBitmap(bmp, new Rectangle(0, 0, panel1.Width, panel1.Height));
            RectangleF bounds = e.PageSettings.PrintableArea;
            float factor = ((float)bmp.Width / (float)bmp.Height);
            e.Graphics.DrawImage(bmp, bounds.Left, bounds.Top, bounds.Height, bounds.Width);
        }

        private void Graphics_Form_SizeChanged(object sender, EventArgs e)
        {
            panel1.Height = this.Height - 179;
            panel1.Width = this.Width - 39;
            chart1.Width = this.Width - 44;

            chart1.Invalidate();
            chart2.Width = this.Width - 44;

            chart2.Top = (int) (panel1.Height * 0.509666);

            if (checkBox1.Checked)
            {
                chart1.Height = (this.Height - 196) / 2;
                chart2.Height = (this.Height - 196) / 2;
            }
            else
            {
                chart1.Height = this.Height - 196;
            }


            chart2.Invalidate();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Graphics_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            GlobalVars.dualPlots = checkBox1.Checked;
        }

        static double GetDouble(string s)
        {
            double d;

            var formatinfo = new System.Globalization.NumberFormatInfo();

            formatinfo.NumberDecimalSeparator = ".";

            if (double.TryParse(s, System.Globalization.NumberStyles.Float, formatinfo, out d))
            {
                return d;
            }

            formatinfo.NumberDecimalSeparator = ",";

            if (double.TryParse(s, System.Globalization.NumberStyles.Float, formatinfo, out d))
            {
                return d;
            }

            throw new SystemException(string.Format("strange number format '{0}'", s));
        }



    }
}
