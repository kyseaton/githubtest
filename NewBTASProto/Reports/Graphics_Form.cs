﻿using System;
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

        //When the type of battery the technology is retrieved store it here
        int technology1 = 0;
        int cell1 = 0;
        int type1 = 0;
        int technology2 = 0;
        int cell2 = 0;
        int type2 = 0;
        Title tt1 = new Title();
        Title tt2 = new Title();

        private readonly object dataBaseLock = new object();

        public Graphics_Form()
        {
            InitializeComponent();
        }

        private void Graphics_Form_Load(object sender, EventArgs e)
        {
            loadWorkOrderLists();

            // Need to set up the title for the charts...
            tt1.Name = "tTitle";
            tt1.Text = "";
            chart1.Titles.Add(tt1);
            tt2.Name = "tTitle";
            tt2.Text = "";
            chart2.Titles.Add(tt2);

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
                    strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                    strAccessSelect = @"SELECT WorkOrderNumber FROM WorkOrders";
                }
                else
                {
                    strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
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
                    MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                    return;
                }
                //  now try to access it
                try
                {
                    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                    lock (dataBaseLock)
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
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
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
                    string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                    string strAccessSelect = @"SELECT StepNumber,TestName, Technology, CustomNoCells FROM Tests WHERE WorkOrderNumber='" + combo1Text + @"'";


                    DataSet testsPerformed = new DataSet();
                    OleDbConnection myAccessConn = null;
                    // try to open the DB
                    try
                    {
                        myAccessConn = new OleDbConnection(strAccessConn);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                        return;
                    }
                    //  now try to access it
                    try
                    {
                        OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                        lock (dataBaseLock)
                        {
                            myAccessConn.Open();
                            myDataAdapter.Fill(testsPerformed, "Tests");
                            myAccessConn.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
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
                        });

                        // update the technology var
                        technology1 = (int)testsPerformed.Tables["Tests"].Rows[1][2];
                        if (Int32.Parse(testsPerformed.Tables["Tests"].Rows[1][3].ToString()) != 0) { cell1 = Int32.Parse(testsPerformed.Tables["Tests"].Rows[1][3].ToString()); }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
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
                    string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                    string strAccessSelect = @"SELECT StepNumber,TestName, Technology, CustomNoCells FROM Tests WHERE WorkOrderNumber='" + combo3Text + @"'";


                    DataSet testsPerformed = new DataSet();
                    OleDbConnection myAccessConn = null;
                    // try to open the DB
                    try
                    {
                        myAccessConn = new OleDbConnection(strAccessConn);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                        return;
                    }
                    //  now try to access it
                    try
                    {
                        OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                        lock (dataBaseLock)
                        {
                            myAccessConn.Open();
                            myDataAdapter.Fill(testsPerformed, "Tests");
                            myAccessConn.Close();
                        }


                    }
                    catch (Exception ex)
                    {   
                        myAccessConn.Close();
                        MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                        return;
                    }

                    try
                    {
                        DataRow emptyRow1 = testsPerformed.Tables["Tests"].NewRow();
                        emptyRow1["TestName"] = "";
                        testsPerformed.Tables["Tests"].Rows.InsertAt(emptyRow1, 0);
                        testsPerformed.Tables["Tests"].Columns.Add("ForList", typeof(string), "StepNumber + ' '+ TestName");

                        this.Invoke((MethodInvoker)delegate()
                        {
                            this.comboBox4.DisplayMember = "ForList";
                            this.comboBox4.ValueMember = "StepNumber";
                            this.comboBox4.DataSource = testsPerformed.Tables["Tests"];

                            comboBox4.Enabled = true;
                        });

                        // update the technology var
                        technology2 = (int)testsPerformed.Tables["Tests"].Rows[1][2];
                        if (Int32.Parse(testsPerformed.Tables["Tests"].Rows[1][3].ToString()) != 0) { cell2 = Int32.Parse(testsPerformed.Tables["Tests"].Rows[1][3].ToString()); }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
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
                    string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                    string strAccessSelect = @"SELECT * FROM ScanData WHERE BWO='" + combo1Text + @"' AND STEP='" + combo2Text.Substring(0, 2) + @"'";

                    //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
                    OleDbConnection myAccessConn = null;
                    // try to open the DB
                    try
                    {
                        myAccessConn = new OleDbConnection(strAccessConn);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                        return;
                    }
                    //  now try to access it
                    try
                    {
                        OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                        lock (dataBaseLock)
                        {
                            myAccessConn.Open();
                            myDataAdapter.Fill(graph1Set, "ScanData");
                            myAccessConn.Close();
                        }


                    }
                    catch (Exception ex)
                    {   
                        myAccessConn.Close();
                        MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
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
                        MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                        return;
                    }
                    //  now try to access it
                    try
                    {
                        OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                        lock (dataBaseLock)
                        {
                            myAccessConn.Open();
                            myDataAdapter.Fill(lookForCable, "Tests");
                            myAccessConn.Close();
                        }
                        

                    }
                    catch (Exception ex)
                    {   myAccessConn.Close();
                        MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                        return;
                    }


                    string cellCable = lookForCable.Tables["Tests"].Rows[0][0].ToString();

                    this.Invoke((MethodInvoker)delegate()
                    {
                        switch (cellCable)
                        {
                            case "1":
                                // update the cells value
                                if (cell1 == 0) { cell1 = 20; }
                                // Battery combobox
                                comboBox5.Items.Clear();
                                comboBox5.Text = "";
                                comboBox5.Items.Add("Voltage");
                                comboBox5.Items.Add("Current");
                                comboBox5.Items.Add("Temperature 1");
                                comboBox5.Items.Add("Temperature 2");
                                comboBox5.Items.Add("Temperature 3");
                                comboBox5.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox6.Items.Clear();
                                comboBox6.Text = "";
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
                            case "3":
                                // update the cells value
                                if (cell1 == 0) { cell1 = 22; }
                                // Battery combobox
                                comboBox5.Items.Clear();
                                comboBox5.Text = "";
                                comboBox5.Items.Add("Voltage 1");
                                comboBox5.Items.Add("Voltage 2");
                                comboBox5.Items.Add("Current");
                                comboBox5.Items.Add("Temperature 1");
                                comboBox5.Items.Add("Temperature 2");
                                comboBox5.Items.Add("Temperature 3");
                                comboBox5.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox6.Items.Clear();
                                comboBox6.Text = "";
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
                                comboBox5.Text = "";
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
                                comboBox6.Text = "";
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
                                // update the cells value
                                if (cell1 == 0) { cell1 = 21; }
                                // Battery combobox
                                comboBox5.Items.Clear();
                                comboBox5.Text = "";
                                comboBox5.Items.Add("Voltage");
                                comboBox5.Items.Add("Current");
                                comboBox5.Items.Add("Temperature 1");
                                comboBox5.Items.Add("Temperature 2");
                                comboBox5.Items.Add("Temperature 3");
                                comboBox5.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox6.Items.Clear();
                                comboBox6.Text = "";
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
                            default:
                                // update the cells value
                                if (cell1 == 0) { cell1 = 20; }
                                // Battery combobox
                                comboBox5.Items.Clear();
                                comboBox5.Text = "";
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
                                comboBox6.Text = "";
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


                        // The final step is to update the type of test that was selected
                        if (comboBox2.Text.Contains("As Recieved")) { type1 = 1; }
                        else if (comboBox2.Text.Contains("Discharge")) { type1 = 2; }
                        else if (comboBox2.Text.Contains("Capacity")) { type1 = 3; }
                        else { type1 = 0; }

                        comboBox5.Enabled = true;
                        comboBox6.Enabled = true;
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
                    string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                    string strAccessSelect = @"SELECT * FROM ScanData WHERE BWO='" + combo3Text + @"' AND STEP='" + combo4Text.Substring(0, 2) + @"'";

                    //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
                    OleDbConnection myAccessConn = null;
                    // try to open the DB
                    try
                    {
                        myAccessConn = new OleDbConnection(strAccessConn);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                        return;
                    }
                    //  now try to access it
                    try
                    {
                        OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                        lock (dataBaseLock)
                        {
                            myAccessConn.Open();
                            myDataAdapter.Fill(graph2Set, "ScanData");
                            myAccessConn.Close();
                        }


                    }
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
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
                        MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                        return;
                    }
                    //  now try to access it
                    try
                    {
                        OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                        lock (dataBaseLock)
                        {
                            myAccessConn.Open();
                            myDataAdapter.Fill(lookForCable, "Tests");
                            myAccessConn.Close();
                        }
 

                    }
                    catch (Exception ex)
                    {
                        myAccessConn.Close();
                        MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                        return;
                    }

                    string cellCable = lookForCable.Tables["Tests"].Rows[0][0].ToString();

                    this.Invoke((MethodInvoker)delegate()
                    {
                        switch (cellCable)
                        {
                            case "1":
                                // update the cells value
                                if (cell2 == 0) { cell2 = 20; }
                                // Battery combobox
                                comboBox7.Items.Clear();
                                comboBox7.Text = "";
                                comboBox7.Items.Add("Voltage");
                                comboBox7.Items.Add("Current");
                                comboBox7.Items.Add("Temperature 1");
                                comboBox7.Items.Add("Temperature 2");
                                comboBox7.Items.Add("Temperature 3");
                                comboBox7.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox8.Items.Clear();
                                comboBox8.Text = "";
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
                            case "3":
                                // update the cells value
                                if (cell2 == 0) { cell2 = 22; }
                                // Battery combobox
                                comboBox7.Items.Clear();
                                comboBox7.Text = "";
                                comboBox7.Items.Add("Voltage 1");
                                comboBox7.Items.Add("Voltage 2");
                                comboBox7.Items.Add("Current");
                                comboBox7.Items.Add("Temperature 1");
                                comboBox7.Items.Add("Temperature 2");
                                comboBox7.Items.Add("Temperature 3");
                                comboBox7.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox8.Items.Clear();
                                comboBox8.Text = "";
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
                                comboBox7.Text = "";
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
                                comboBox8.Text = "";
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
                                // update the cells value
                                if (cell2 == 0) { cell2 = 21; }
                                // Battery combobox
                                comboBox7.Items.Clear();
                                comboBox7.Text = "";
                                comboBox7.Items.Add("Voltage");
                                comboBox7.Items.Add("Current");
                                comboBox7.Items.Add("Temperature 1");
                                comboBox7.Items.Add("Temperature 2");
                                comboBox7.Items.Add("Temperature 3");
                                comboBox7.Items.Add("Temperature 4");
                                // Cells combobox
                                comboBox8.Items.Clear();
                                comboBox8.Text = "";
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
                            default:
                                // update the cells value
                                if (cell2 == 0) { cell2 = 20; }
                                // Battery combobox
                                comboBox7.Items.Clear();
                                comboBox7.Text = "";
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
                                comboBox8.Text = "";
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



                        // The final step is to update the type of test that was selected
                        if (comboBox4.Text.Contains("As Recieved")) { type2 = 1; }
                        else if (comboBox4.Text.Contains("Discharge")) { type2 = 2; }
                        else if (comboBox4.Text.Contains("Capacity")) { type2 = 3; }
                        else { type2 = 0; }

                        comboBox7.Enabled = true;
                        comboBox8.Enabled = true;
                    });// end invoke
                });// end helper thread
            }
        }

        private void comboBox5_SelectedValueChanged(object sender, EventArgs e)
        {
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
                        q = 8;
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

                for (int i = 0; i < graph1Set.Tables[0].Rows.Count; i++)
                {
                    series1.Points.AddXY((int)(double.Parse(graph1Set.Tables[0].Rows[i][7].ToString()) * 1440), graph1Set.Tables[0].Rows[i][q]);
                    // color test
                    series1.Points[i].Color = pointColor(technology1, cell1, double.Parse(graph1Set.Tables[0].Rows[i][q].ToString()), type1);
                }

                tt1.Text = "Work Order:  " + comboBox1.Text + "     Test:  " + comboBox2.Text + "     Date:  " + graph1Set.Tables[0].Rows[0][4].ToString();
                chart1.Invalidate();
                chart1.ChartAreas[0].RecalculateAxesScale();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
            }



        }
        private void comboBox6_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                int q;
                // only do something if the radio button is selected
                if (radioButton2.Checked == false || comboBox6.SelectedIndex < 0) { return; }
                // Here we will look at the Value selected and then plot graph1Set

                //find out which graph to plot from the selected text
                switch (comboBox6.Text)
                {
                    case "Ending Voltages":
                        q = 999;
                        break;
                    case "Cell 1":
                        q = 14;
                        break;
                    case "Cell 2":
                        q = 15;
                        break;
                    case "Cell 3":
                        q = 16;
                        break;
                    case "Cell 4":
                        q = 17;
                        break;
                    case "Cell 5":
                        q = 18;
                        break;
                    case "Cell 6":
                        q = 19;
                        break;
                    case "Cell 7":
                        q = 20;
                        break;
                    case "Cell 8":
                        q = 21;
                        break;
                    case "Cell 9":
                        q = 22;
                        break;
                    case "Cell 10":
                        q = 23;
                        break;
                    case "Cell 11":
                        q = 24;
                        break;
                    case "Cell 12":
                        q = 25;
                        break;
                    case "Cell 13":
                        q = 26;
                        break;
                    case "Cell 14":
                        q = 27;
                        break;
                    case "Cell 15":
                        q = 28;
                        break;
                    case "Cell 16":
                        q = 29;
                        break;
                    case "Cell 17":
                        q = 30;
                        break;
                    case "Cell 18":
                        q = 31;
                        break;
                    case "Cell 19":
                        q = 32;
                        break;
                    case "Cell 20":
                        q = 33;
                        break;
                    case "Cell 21":
                        q = 34;
                        break;
                    case "Cell 22":
                        q = 35;
                        break;
                    case "Cell 23":
                        q = 36;
                        break;
                    case "Cell 24":
                        q = 37;
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
                        series1.Points.AddXY(i + 1, graph1Set.Tables[0].Rows[graph1Set.Tables[0].Rows.Count - 1][i + 14]);
                        // color test
                        series1.Points[i].Color = pointColor(technology1, 1, double.Parse(graph1Set.Tables[0].Rows[graph1Set.Tables[0].Rows.Count - 1][i + 14].ToString()), type1);
                    }
                }
                else
                {
                    for (int i = 0; i < graph1Set.Tables[0].Rows.Count; i++)
                    {
                        series1.Points.AddXY((int)(double.Parse(graph1Set.Tables[0].Rows[i][7].ToString()) * 1440), graph1Set.Tables[0].Rows[i][q]);
                        // color test
                        series1.Points[i].Color = pointColor(technology1, 1, double.Parse(graph1Set.Tables[0].Rows[i][q].ToString()), type1);
                    }
                }

                tt1.Text = "Work Order:  " + comboBox1.Text + "     Test:  " + comboBox2.Text + "     Date:  " + graph1Set.Tables[0].Rows[0][4].ToString();
                chart1.Invalidate();
                chart1.ChartAreas[0].RecalculateAxesScale();


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
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
        private Color pointColor(int tech, int Cells, double Value, int type)
        {
            if (radioButton1.Checked == true)
            {
                switch (comboBox5.Text)
                {
                    case "Current":
                        return System.Drawing.Color.Blue;
                    case "Temperature 1":
                    case "Temperature 2":
                    case "Temperature 3":
                    case "Temperature 4":
                        return System.Drawing.Color.Orange;
                    default:
                        break;
                }
            }

            // normal vented NiCds
            double Min1 = 0.25;
            double Min2 = 1.5;
            double Min3 = 1.55;
            double Min4 = 1.7;
            double Max = 1.75;

            // special case for cable 10, sealed NiCds
            if (tech == 1)
            {
                Min1 = 0.25;
                Min2 = 1.45;
                Min3 = 1.5;
                Min4 = 1.65;
                Max = 1.7;
            }

            // these are the discharging cases...
            switch (type)
            {
                
                // these are the As Recieved color setting    
                case 1:
                    Min1 = 0.1;
                    Min2 = 1.2;
                    Min3 = 1.25;
                    Min4 = 1.25;
                    break;
                // these are the Discharge settings
                case 2:
                    Min1 = 0;
                    Min2 = 0.5;
                    Min3 = 0.5;
                    Min4 = 0.5;
                    break;
                // these are the Capacity
                case 3:
                    Min1 = 1;
                    Min2 = 1;
                    Min3 = 1.05;
                    Min4 = 11.7;
                    Max = 1.25;
                    break;
                default:
                    break;
            }

            // scale the limits for the number of cells in the battery
            Min1 *= Cells;
            Min2 *= Cells;
            Min3 *= Cells;
            Min4 *= Cells;
            Max *= Cells;



            // with all of that said, let's start picking colors!
            // this is for all charging operations not involving lead acid
            if (tech != 2 && type == 0)
            {
                if (Value < Min2) { return System.Drawing.Color.Yellow; }
                else if (Value >= Min2 && Value < Min3) { return System.Drawing.Color.Orange; }
                else if (Value >= Min3 && Value < Min4) { return System.Drawing.Color.Green; }
                else if (Value >= Min4 && Value < Max) { return System.Drawing.Color.Blue; }
                else { return System.Drawing.Color.Red; }
            }
            // lead acid case
            else if(type == 0)
            {
                return System.Drawing.Color.Orange;
            }
            // this is for the Capacity test, Discharge and As Recieved
            else if (tech != 2)
            {
                if (Value < Min1) { return System.Drawing.Color.Red; }
                else if (Value < Min2) { return System.Drawing.Color.Yellow; }
                else if (Value < Min3) { return System.Drawing.Color.Orange; }
                else if (Value < Min4) { return System.Drawing.Color.Green; }
                else if (Value < Min2) { return System.Drawing.Color.Orange; }
            }
            //Finally the lead acid Capacity test, Discharge and As Recieved case
            else
            {
                if (Value > Cells * 1.75) { return System.Drawing.Color.Green; }
                else if (Value >= Cells * 1.67) { return System.Drawing.Color.Orange; }
                return System.Drawing.Color.Red; 
            }

            return System.Drawing.Color.Green;


        }
        private Color pointColor2(int tech, int Cells, double Value, int type)
        {
            if (radioButton3.Checked == true)
            {
                switch (comboBox7.Text)
                {
                    case "Current":
                        return System.Drawing.Color.Blue;
                    case "Temperature 1":
                    case "Temperature 2":
                    case "Temperature 3":
                    case "Temperature 4":
                        return System.Drawing.Color.Orange;
                    default:
                        break;
                }
            }

            // normal vented NiCds
            double Min1 = 0.25;
            double Min2 = 1.5;
            double Min3 = 1.55;
            double Min4 = 1.7;
            double Max = 1.75;

            // special case for cable 10, sealed NiCds
            if (tech == 1)
            {
                Min1 = 0.25;
                Min2 = 1.45;
                Min3 = 1.5;
                Min4 = 1.65;
                Max = 1.7;
            }

            // these are the discharging cases...
            switch (type)
            {

                // these are the As Recieved color setting    
                case 1:
                    Min1 = 0.1;
                    Min2 = 1.2;
                    Min3 = 1.25;
                    Min4 = 1.25;
                    break;
                // these are the Discharge settings
                case 2:
                    Min1 = 0;
                    Min2 = 0.5;
                    Min3 = 0.5;
                    Min4 = 0.5;
                    break;
                // these are the Capacity
                case 3:
                    Min1 = 1;
                    Min2 = 1;
                    Min3 = 1.05;
                    Min4 = 11.7;
                    Max = 1.25;
                    break;
                default:
                    break;
            }

            // scale the limits for the number of cells in the battery
            Min1 *= Cells;
            Min2 *= Cells;
            Min3 *= Cells;
            Min4 *= Cells;
            Max *= Cells;



            // with all of that said, let's start picking colors!
            // this is for all charging operations not involving lead acid
            if (tech != 2 && type == 0)
            {
                if (Value < Min2) { return System.Drawing.Color.Yellow; }
                else if (Value >= Min2 && Value < Min3) { return System.Drawing.Color.Orange; }
                else if (Value >= Min3 && Value < Min4) { return System.Drawing.Color.Green; }
                else if (Value >= Min4 && Value < Max) { return System.Drawing.Color.Blue; }
                else { return System.Drawing.Color.Red; }
            }
            // lead acid case
            else if (type == 0)
            {
                return System.Drawing.Color.Orange;
            }
            // this is for the Capacity test, Discharge and As Recieved
            else if (tech != 2)
            {
                if (Value < Min1) { return System.Drawing.Color.Red; }
                else if (Value < Min2) { return System.Drawing.Color.Yellow; }
                else if (Value < Min3) { return System.Drawing.Color.Orange; }
                else if (Value < Min4) { return System.Drawing.Color.Green; }
                else if (Value < Min2) { return System.Drawing.Color.Orange; }
            }
            //Finally the lead acid Capacity test, Discharge and As Recieved case
            else
            {
                if (Value > Cells * 1.75) { return System.Drawing.Color.Green; }
                else if (Value >= Cells * 1.67) { return System.Drawing.Color.Orange; }
                return System.Drawing.Color.Red;
            }

            return System.Drawing.Color.Green;


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
                        q = 8;
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

                for (int i = 0; i < graph2Set.Tables[0].Rows.Count; i++)
                {
                    series2.Points.AddXY((int)(double.Parse(graph2Set.Tables[0].Rows[i][7].ToString()) * 1440), graph2Set.Tables[0].Rows[i][q]);
                    // color test
                    series2.Points[i].Color = pointColor2(technology2, cell2, double.Parse(graph2Set.Tables[0].Rows[i][q].ToString()), type2);
                }
                tt2.Text = "Work Order:  " + comboBox3.Text + "     Test:  " + comboBox4.Text + "     Date:  " + graph2Set.Tables[0].Rows[0][4].ToString();
                chart2.Invalidate();
                chart2.ChartAreas[0].RecalculateAxesScale();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
            }

        }

        private void comboBox8_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                int q;
                // only do something if the radio button is selected
                if (radioButton4.Checked == false || comboBox8.SelectedIndex < 0) { return; }
                // Here we will look at the Value selected and then plot graph1Set

                //find out which graph to plot from the selected text
                switch (comboBox8.Text)
                {
                    case "Ending Voltages":
                        q = 999;
                        break;
                    case "Cell 1":
                        q = 14;
                        break;
                    case "Cell 2":
                        q = 15;
                        break;
                    case "Cell 3":
                        q = 16;
                        break;
                    case "Cell 4":
                        q = 17;
                        break;
                    case "Cell 5":
                        q = 18;
                        break;
                    case "Cell 6":
                        q = 19;
                        break;
                    case "Cell 7":
                        q = 20;
                        break;
                    case "Cell 8":
                        q = 21;
                        break;
                    case "Cell 9":
                        q = 22;
                        break;
                    case "Cell 10":
                        q = 23;
                        break;
                    case "Cell 11":
                        q = 24;
                        break;
                    case "Cell 12":
                        q = 25;
                        break;
                    case "Cell 13":
                        q = 26;
                        break;
                    case "Cell 14":
                        q = 27;
                        break;
                    case "Cell 15":
                        q = 28;
                        break;
                    case "Cell 16":
                        q = 29;
                        break;
                    case "Cell 17":
                        q = 30;
                        break;
                    case "Cell 18":
                        q = 31;
                        break;
                    case "Cell 19":
                        q = 32;
                        break;
                    case "Cell 20":
                        q = 33;
                        break;
                    case "Cell 21":
                        q = 34;
                        break;
                    case "Cell 22":
                        q = 35;
                        break;
                    case "Cell 23":
                        q = 36;
                        break;
                    case "Cell 24":
                        q = 37;
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
                    for (int i = 0; i < cell1; i++)
                    {
                        series2.Points.AddXY(i + 1, graph2Set.Tables[0].Rows[graph2Set.Tables[0].Rows.Count - 1][i + 14]);
                        // color test
                        series2.Points[i].Color = pointColor2(technology2, 1, double.Parse(graph2Set.Tables[0].Rows[graph2Set.Tables[0].Rows.Count - 1][i + 14].ToString()), type2);
                    }
                }
                else
                {
                    for (int i = 0; i < graph2Set.Tables[0].Rows.Count; i++)
                    {
                        series2.Points.AddXY((int)(double.Parse(graph2Set.Tables[0].Rows[i][7].ToString()) * 1440), graph2Set.Tables[0].Rows[i][q]);
                        // color test
                        series2.Points[i].Color = pointColor2(technology2, 1, double.Parse(graph2Set.Tables[0].Rows[i][q].ToString()), type2);
                    }
                }

                tt2.Text = "Work Order:  " + comboBox3.Text + "     Test:  " + comboBox4.Text + "     Date:  " + graph2Set.Tables[0].Rows[0][4].ToString();
                chart2.Invalidate();
                chart2.ChartAreas[0].RecalculateAxesScale();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
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
            chart1.Height = (this.Height - 196)/2;
            chart1.Invalidate();
            chart2.Width = this.Width - 44;
            chart2.Height = (this.Height - 196) / 2;
            chart2.Top = (int) (panel1.Height * 0.509666);
            chart2.Invalidate();
        }    



    }
}
