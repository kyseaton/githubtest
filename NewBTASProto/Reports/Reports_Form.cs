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
using Microsoft.Reporting.WinForms;

namespace NewBTASProto
{
    public partial class Reports_Form : Form
    {
        //form wide dataSets used to fill in the two graphs
        DataSet reportSet = new DataSet();
        DataSet testsPerformed = new DataSet();


        public Reports_Form()
        {
            InitializeComponent();
        }

        private void Reports_Form_Load(object sender, EventArgs e)
        {

            /*************************Load Global data into MetaData data Table ************************/

            // create datatable
            DataTable MetaDT = new DataTable("MetaData");

            // add columns
            MetaDT.Columns.Add("gBusinessName", typeof(string));
            MetaDT.Columns.Add("gUseF", typeof(string));
            MetaDT.Columns.Add("gPos2Neg", typeof(string));

            // insert data rows
            MetaDT.Rows.Add(GlobalVars.businessName, GlobalVars.useF, GlobalVars.Pos2Neg);


            // bind datatable to report viewer
            this.reportViewer.Reset();
            this.reportViewer.ProcessingMode = ProcessingMode.Local;
            this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.Report1.rdlc";
            this.reportViewer.LocalReport.DataSources.Clear();
            this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("MetaData", MetaDT));
            this.reportViewer.RefreshReport();

            /*********************************************************/

            loadWorkOrderLists();
            this.reportViewer.RefreshReport();
        }


        private void loadWorkOrderLists()
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

                myAccessConn.Open();
                myDataAdapter.Fill(workOrderList1, "ScanData");
                myDataAdapter.Fill(workOrderList2, "ScanData");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                return;
            }
            finally
            {
                myAccessConn.Close();
            }

            DataRow emptyRow1 = workOrderList1.Tables["ScanData"].NewRow();
            emptyRow1["WorkOrderNumber"] = "";
            workOrderList1.Tables["ScanData"].Rows.InsertAt(emptyRow1,0);

            this.comboBox1.DisplayMember = "WorkOrderNumber";
            this.comboBox1.ValueMember = "WorkOrderNumber";
            this.comboBox1.DataSource = workOrderList1.Tables["ScanData"];

            // remember to clear everything!
            this.comboBox2.DataSource = null;
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            // reset the test combo box
            comboBox3.SelectedIndex = -1;
            testsPerformed = new DataSet();

            if (comboBox1.SelectedIndex <= 0) 
            {
                comboBox2.DataSource = null;
                return; 
            }
            else
            {
                // Open database containing all the battery data....
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                string strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + comboBox1.Text + @"'";

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

                    myAccessConn.Open();
                    myDataAdapter.Fill(testsPerformed, "Tests");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    return;
                }
                finally
                {
                    myAccessConn.Close();
                }

                try
                {
                    DataRow emptyRow1 = testsPerformed.Tables["Tests"].NewRow();
                    emptyRow1["TestName"] = "";
                    testsPerformed.Tables["Tests"].Rows.InsertAt(emptyRow1, 0);
                    testsPerformed.Tables["Tests"].Columns.Add("ForList", typeof(string), "StepNumber + ' - '+ TestName");

                    this.comboBox2.DisplayMember = "ForList";
                    this.comboBox2.ValueMember = "StepNumber";
                    this.comboBox2.DataSource = testsPerformed.Tables["Tests"];

                }
                catch(Exception ex)
                {
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    return;
                }
                

            }
        }


        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            // Clear the report type selection
            comboBox3.SelectedIndex = -1;
            // load data in the report specific methods below
        }


        private void comboBox3_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                switch(comboBox3.SelectedIndex)
                {
                    case 0:
                        batReport();
                        break;
                    case 1:
                        cellData();
                        break;
                    case 2:
                        testSummary();
                        break;
                    case 3:
                        workOrderLog();
                        break;
                    case 4:
                        workOrderSummary();
                        break;
                    default:
                        break;
                }
            }

            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
         
        }

        private void workOrderSummary()
        {
            reportSet = new  DataSet();

            if (!(comboBox1.SelectedIndex < 1)) // make sure we have a selection to act on...
            {
                // FIRST CLEAR THE OLD DATA SET!
                reportSet.Clear();
                // Open database containing all the battery data....
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                string strAccessSelect = @"SELECT DATE FROM ScanData WHERE BWO='" + comboBox1.Text + "'";

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

                    myAccessConn.Open();
                    myDataAdapter.Fill(reportSet, "ScanData");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    return;
                }
                finally
                {
                    myAccessConn.Close();
                }
            }

            DataTable testTable = new DataTable();


            testTable = testsPerformed.Tables[0].Copy();
            testTable.Rows[0].Delete();

            if (!(comboBox1.SelectedIndex < 0)) // make sure we have a selection to act on...
            {

                // We have the data in testsPerformed, so lets pass it over to the matching report

                // bind datatable to report viewer
                this.reportViewer.Reset();
                this.reportViewer.ProcessingMode = ProcessingMode.Local;
                this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WorkOrderSummary.rdlc";
                this.reportViewer.LocalReport.DataSources.Clear();

                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("reportSet", reportSet.Tables[0]));

                /*************************Load testsPerformed into Tests  ************************/


                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("Tests", testTable));

                /*************************Load Global data into MetaData data Table ***********************
                // create datatable
                DataTable MetaDT = new DataTable("MetaData");
                MetaDT.Clear();

                // add columns
                MetaDT.Columns.Add("gBusinessName", typeof(string));
                MetaDT.Columns.Add("gUseF", typeof(string));
                MetaDT.Columns.Add("gPos2Neg", typeof(string));
                MetaDT.Columns.Add("testComboName", typeof(string));
                MetaDT.Columns.Add("cellsCable", typeof(string));
                MetaDT.Columns.Add("shuntCable", typeof(string));
                MetaDT.Columns.Add("tempCable", typeof(string));

                // insert data rows
                MetaDT.Rows.Add(GlobalVars.businessName, GlobalVars.useF, GlobalVars.Pos2Neg, comboBox2.Text, testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][15].ToString(), testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][16].ToString(), testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][17].ToString());
                */

                DataTable MetaDT = new DataTable("MetaData");
                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("MetaData", MetaDT));

                this.reportViewer.RefreshReport();

                /*********************************************************/

                // finally enable the reportview
                reportViewer.Enabled = true;

            }

        }

        private void workOrderLog()
        {
            reportSet = new DataSet();

            if (!(comboBox1.SelectedIndex < 1)) // make sure we have a selection to act on...
            {
                // FIRST CLEAR THE OLD DATA SET!
                reportSet.Clear();
                // Open database containing all the battery data....
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                string strAccessSelect = @"SELECT DATE FROM ScanData WHERE BWO='" + comboBox1.Text +"'";

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

                    myAccessConn.Open();
                    myDataAdapter.Fill(reportSet, "ScanData");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    return;
                }
                finally
                {
                    myAccessConn.Close();
                }
            }

            DataTable testTable = new DataTable();

            
            testTable = testsPerformed.Tables[0].Copy();
            testTable.Rows[0].Delete();

            if (!(comboBox1.SelectedIndex < 0)) // make sure we have a selection to act on...
            {
              
                // We have the data in testsPerformed, so lets pass it over to the matching report

                // bind datatable to report viewer
                this.reportViewer.Reset();
                this.reportViewer.ProcessingMode = ProcessingMode.Local;
                this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WorkOrderLog.rdlc"; 
                this.reportViewer.LocalReport.DataSources.Clear();

                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("reportSet", reportSet.Tables[0]));

                /*************************Load testsPerformed into Tests  ************************/


                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("Tests", testTable));

                /*************************Load Global data into MetaData data Table ***********************
                // create datatable
                DataTable MetaDT = new DataTable("MetaData");
                MetaDT.Clear();

                // add columns
                MetaDT.Columns.Add("gBusinessName", typeof(string));
                MetaDT.Columns.Add("gUseF", typeof(string));
                MetaDT.Columns.Add("gPos2Neg", typeof(string));
                MetaDT.Columns.Add("testComboName", typeof(string));
                MetaDT.Columns.Add("cellsCable", typeof(string));
                MetaDT.Columns.Add("shuntCable", typeof(string));
                MetaDT.Columns.Add("tempCable", typeof(string));

                // insert data rows
                MetaDT.Rows.Add(GlobalVars.businessName, GlobalVars.useF, GlobalVars.Pos2Neg, comboBox2.Text, testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][15].ToString(), testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][16].ToString(), testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][17].ToString());
                */

                DataTable MetaDT = new DataTable("MetaData");
                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("MetaData", MetaDT));

                this.reportViewer.RefreshReport();

                /*********************************************************/

                // finally enable the reportview
                reportViewer.Enabled = true;

            }
        }

        private void testSummary()
        {
            reportSet = new DataSet();

            if (!(comboBox2.SelectedIndex < 1)) // make sure we have a selection to act on...
            {
                // FIRST CLEAR THE OLD DATA SET!
                reportSet.Clear();
                // Open database containing all the battery data....
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                string strAccessSelect = @"SELECT TOP 1 DATE,RDG,ETIME,CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24 FROM ScanData WHERE BWO='" + comboBox1.Text + @"' AND STEP='" + comboBox2.Text.Substring(0, 2) + @"' ORDER BY DATE DESC";

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

                    myAccessConn.Open();
                    myDataAdapter.Fill(reportSet, "ScanData");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    return;
                }
                finally
                {
                    myAccessConn.Close();
                }
                //  Here is where we do the Cell voltage PASS/FAIL determinations

                bool charge = false;        // default to discharge

                if (testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][5].ToString() == "Full Charge-6" ||
                    testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][5].ToString() == "Full Charge-4" ||
                    testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][5].ToString() == "Top Charge-4" ||
                    testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][5].ToString() == "Top Charge-2" ||
                    testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][5].ToString() == "Top Charge-1" ||
                    testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][5].ToString() == "Slow Charge" ||
                    testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][5].ToString() == "Custom Chg #1" ||
                    testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][5].ToString() == "Custom Chg #2" ||
                    testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][5].ToString() == "Custom Chg #3" ||
                    testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][5].ToString() == "Reflex Chg-1" ||
                    testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][5].ToString() == "Custom Chg") 
                    { charge = true; }

                // now run tests on final cell voltages and add to reportSet
                reportSet.Tables[0].Rows.Add();
                for (int i = 3; i < 27; i++)
                {
                    if (testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][15].ToString() == "10")
                    {
                        reportSet.Tables[0].Rows[1][i] = "No Data"; 
                    }
                    else if (charge)
                    {
                        if (double.Parse(reportSet.Tables[0].Rows[0][i].ToString()) > 1.5 && double.Parse(reportSet.Tables[0].Rows[0][i].ToString()) <1.75) {reportSet.Tables[0].Rows[1][i] = "OK";}
                        else if (double.Parse(reportSet.Tables[0].Rows[0][i].ToString()) > 1.75) { reportSet.Tables[0].Rows[1][i] = "FAIL! Overvoltage!"; }
                        else { reportSet.Tables[0].Rows[1][i] = "FAIL!"; }
                    }
                    else
                    {
                        if (double.Parse(reportSet.Tables[0].Rows[0][i].ToString()) > 1) { reportSet.Tables[0].Rows[1][i] = "OK"; }
                        else { reportSet.Tables[0].Rows[1][i] = "FAIL!"; }
                    }
                }  // end for



                // Now that we have the data in reportSet along with testsPerformed lets pass it over to the matching report
                /*************************Load reportSet into reportSet  ************************/

                // bind datatable to report viewer
                this.reportViewer.Reset();
                this.reportViewer.ProcessingMode = ProcessingMode.Local;
                

               switch (testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][15].ToString())
                {
                    case "3":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.TestSummaryPN22.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.TestSummaryNP22.rdlc"; }
                        break;
                    case "4":
                    case "21":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.TestSummaryPN21.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.TestSummaryNP21.rdlc"; }
                        break;
                    default:
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.TestSummaryPN20.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.TestSummaryNP20.rdlc"; }
                        break;
                }// end switch

                this.reportViewer.LocalReport.DataSources.Clear();



                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("reportSet", reportSet.Tables[0]));

                /*************************Load testsPerformed into Tests  ************************/

                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("Tests", testsPerformed.Tables[0]));

                /*************************Load Global data into MetaData data Table ************************/
                // create datatable
                DataTable MetaDT = new DataTable("MetaData");

                // add columns
                MetaDT.Columns.Add("gBusinessName", typeof(string));
                MetaDT.Columns.Add("gUseF", typeof(string));
                MetaDT.Columns.Add("gPos2Neg", typeof(string));
                MetaDT.Columns.Add("testComboName", typeof(string));
                MetaDT.Columns.Add("cellsCable", typeof(string));
                MetaDT.Columns.Add("shuntCable", typeof(string));
                MetaDT.Columns.Add("tempCable", typeof(string));

                // insert data rows
                MetaDT.Rows.Add(GlobalVars.businessName, GlobalVars.useF, GlobalVars.Pos2Neg, comboBox2.Text, testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][15].ToString(), testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][16].ToString(), testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][17].ToString());

                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("MetaData", MetaDT));
                this.reportViewer.RefreshReport();

                /*********************************************************/

                // finally enable the reportview
                reportViewer.Enabled = true;

            }
        }

        private void cellData()
        {

            if (!(comboBox2.SelectedIndex < 1)) // make sure we have a selection to act on...
            {
                // FIRST CLEAR THE OLD DATA SET!
                reportSet.Clear();
                // Open database containing all the battery data....
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                string strAccessSelect = @"SELECT DATE,RDG,ETIME,CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24 FROM ScanData WHERE BWO='" + comboBox1.Text + @"' AND STEP='" + comboBox2.Text.Substring(0, 2) + @"'";

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

                    myAccessConn.Open();
                    myDataAdapter.Fill(reportSet, "ScanData");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    return;
                }
                finally
                {
                    myAccessConn.Close();
                }


                // Now that we have the data in reportSet along with testsPerformed lets pass it over to the matching report
                /*************************Load reportSet into reportSet  ************************/

                // bind datatable to report viewer
                this.reportViewer.Reset();
                this.reportViewer.ProcessingMode = ProcessingMode.Local;
                switch (testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][15].ToString())
                {
                    case "1":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN20.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP20.rdlc"; }
                        break;
                    case "3":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN22.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP22.rdlc"; }
                        break;
                    case "4":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN21.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP21.rdlc"; }
                        break;
                    case "21":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN21.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP21.rdlc"; }
                        break;
                    default:
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN20.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP20.rdlc"; }
                        break;
                }// end switch

                this.reportViewer.LocalReport.DataSources.Clear();



                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("reportSet", reportSet.Tables[0]));

                /*************************Load testsPerformed into Tests  ************************/

                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("Tests", testsPerformed.Tables[0]));

                /*************************Load Global data into MetaData data Table ************************/
                // create datatable
                DataTable MetaDT = new DataTable("MetaData");

                // add columns
                MetaDT.Columns.Add("gBusinessName", typeof(string));
                MetaDT.Columns.Add("gUseF", typeof(string));
                MetaDT.Columns.Add("gPos2Neg", typeof(string));
                MetaDT.Columns.Add("testComboName", typeof(string));
                MetaDT.Columns.Add("cellsCable", typeof(string));
                MetaDT.Columns.Add("shuntCable", typeof(string));
                MetaDT.Columns.Add("tempCable", typeof(string));

                // insert data rows
                MetaDT.Rows.Add(GlobalVars.businessName, GlobalVars.useF, GlobalVars.Pos2Neg, comboBox2.Text, testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][15].ToString(), testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][16].ToString(), testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][17].ToString());

                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("MetaData", MetaDT));
                this.reportViewer.RefreshReport();

                /*********************************************************/

                // finally enable the reportview
                reportViewer.Enabled = true;

            }
        }

        private void batReport()
        {
            if (!(comboBox2.SelectedIndex < 1)) // make sure we have a selection to act on...
            {
                // FIRST CLEAR THE OLD DATA SET!
                reportSet.Clear();
                // Open database containing all the battery data....
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                string strAccessSelect = @"SELECT RDG,DATE,ETIME,CUR1,VB1,VB2,VB3,VB4,BT1,BT2,BT3,BT4,REF,BT5 FROM ScanData WHERE BWO='" + comboBox1.Text + @"' AND STEP='" + comboBox2.Text.Substring(0, 2) + @"'";

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

                    myAccessConn.Open();
                    myDataAdapter.Fill(reportSet, "ScanData");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    return;
                }
                finally
                {
                    myAccessConn.Close();
                }


                // Now that we have the data in reportSet along with testsPerformed lets pass it over to the matching report
                /*************************Load reportSet into reportSet  ************************/

                // bind datatable to report viewer
                this.reportViewer.Reset();
                this.reportViewer.ProcessingMode = ProcessingMode.Local;

                switch (testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][15].ToString())
                {
                    case "10":
                        this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.BatteryData4.rdlc";
                        break;
                    case "20":
                    case "4":
                        this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.BatteryData3.rdlc";
                        break;
                    case "19":
                    case "3":
                        this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.BatteryData2.rdlc";
                        break;
                    default:
                        this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.BatteryData1.rdlc";
                        break;
                }// end switch

                this.reportViewer.LocalReport.DataSources.Clear();
                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("reportSet", reportSet.Tables[0]));

                /*************************Load testsPerformed into Tests  ************************/

                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("Tests", testsPerformed.Tables[0]));

                /*************************Load Global data into MetaData data Table ************************/
                // create datatable
                DataTable MetaDT = new DataTable("MetaData");

                // add columns
                MetaDT.Columns.Add("gBusinessName", typeof(string));
                MetaDT.Columns.Add("gUseF", typeof(string));
                MetaDT.Columns.Add("gPos2Neg", typeof(string));
                MetaDT.Columns.Add("testComboName", typeof(string));
                MetaDT.Columns.Add("cellsCable", typeof(string));
                MetaDT.Columns.Add("shuntCable", typeof(string));
                MetaDT.Columns.Add("tempCable", typeof(string));


                // insert data rows
                MetaDT.Rows.Add(GlobalVars.businessName, GlobalVars.useF, GlobalVars.Pos2Neg, comboBox2.Text, testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][15].ToString(), testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][16].ToString(), testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][17].ToString());

                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("MetaData", MetaDT));
                this.reportViewer.RefreshReport();

                /*********************************************************/

                // finally enable the reportview
                reportViewer.Enabled = true;

            }

        }


        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            loadWorkOrderLists();
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
           // Bitmap bmp = new Bitmap(panel1.Width, panel1.Height, panel1.CreateGraphics());
            //panel1.DrawToBitmap(bmp, new Rectangle(0, 0, panel1.Width, panel1.Height));
            //RectangleF bounds = e.PageSettings.PrintableArea;
            //float factor = ((float)bmp.Width / (float)bmp.Height);
            //e.Graphics.DrawImage(bmp, bounds.Left, bounds.Top, bounds.Height, bounds.Width);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private static double ConvertCelsiusToFahrenheit(double c)
        {
            return ((9.0 / 5.0) * c) + 32;
        }

        private void reportViewer1_Load(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Reports_Form_SizeChanged(object sender, EventArgs e)
        {
            reportViewer.Width = this.Width - 43;
            reportViewer.Height = this.Height - 106;
        }


    }
}
