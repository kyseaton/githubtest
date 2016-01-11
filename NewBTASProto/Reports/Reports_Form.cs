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
        //form wide dataSets used to fill in the report
        DataSet reportSet = new DataSet();
        DataSet testsPerformed = new DataSet();

        string curStep = "";
        string curWorkOrder = "";
        bool startup = true;
        
        //        ReportParameter Path = new ReportParameter("ImagePath", filePath.A;
        
        //Me.reportViewer1.LocalReport.SetParameters(New ReportParameter() {Path})


        public Reports_Form(string currentStep = "", string currentWorkOrder = "")
        {
            InitializeComponent();

            curStep = currentStep;
            curWorkOrder = currentWorkOrder;

            this.reportViewer.LocalReport.SubreportProcessing += new SubreportProcessingEventHandler(MySubreportEventHandler);
        }

        void MySubreportEventHandler(object sender, SubreportProcessingEventArgs e)
        {
            e.DataSources.Add(new ReportDataSource("Data", reportSet));
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

            this.reportViewer.LocalReport.EnableExternalImages = true;
            ReportParameter parameter = new ReportParameter("Path", "file:////" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\rp_logo.jpg");
            this.reportViewer.LocalReport.SetParameters(parameter);

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
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                strAccessSelect = @"SELECT WorkOrderNumber FROM WorkOrders";
            }
            else
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
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
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
            }

            DataRow emptyRow1 = workOrderList1.Tables["ScanData"].NewRow();
            emptyRow1["WorkOrderNumber"] = "";
            workOrderList1.Tables["ScanData"].Rows.InsertAt(emptyRow1,0);

            this.comboBox1.DisplayMember = "WorkOrderNumber";
            this.comboBox1.ValueMember = "WorkOrderNumber";
            this.comboBox1.DataSource = workOrderList1.Tables["ScanData"];

            // remember to clear everything!
            this.comboBox2.DataSource = null;

            if (startup)
            {
                //Now set the comboboxes to the current station and workorder...

                // we need to split up the work orders if we have multiple work orders on a single line...
                string tempWOS = curWorkOrder;
                char[] delims = { ' ' };
                string[] A = tempWOS.Split(delims);
                curWorkOrder = A[0];

                comboBox1.Text = curWorkOrder.Trim();
                //comboBox1_SelectedValueChanged(this, null);


                startup = false;
                if (curStep.Length < 2)
                {
                    curStep = "0" + curStep;
                }
                try
                {
                    comboBox2.SelectedIndex = comboBox2.FindString(curStep);
                }
                catch
                {
                    // do nothing!
                }
                

                if (curStep != "" && curWorkOrder != "")
                {
                    comboBox3.SelectedIndex = 0;
                }
            }
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
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                string strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + comboBox1.Text + @"' ORDER BY StepNumber ASC";

                OleDbConnection myAccessConn = null;
                // try to open the DB
                try
                {
                    myAccessConn = new OleDbConnection(strAccessConn);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    
                }

                // now we need to see if there were water level values recorded...
                DataSet WaterLevel = new DataSet();
                strAccessSelect = @"SELECT * FROM WaterLevel WHERE WorkOrderNumber='" + comboBox1.Text + @"' ORDER BY WLID ASC";
                
                try
                {
                    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                    lock (Main_Form.dataBaseLock)
                    {
                        myAccessConn.Open();
                        myDataAdapter.Fill(WaterLevel, "WaterLevel");
                        myAccessConn.Close();
                    }

                }
                catch
                {
                    // do nothing...
                }


                try
                {
                    DataRow emptyRow1 = testsPerformed.Tables["Tests"].NewRow();
                    emptyRow1["TestName"] = "";
                    testsPerformed.Tables["Tests"].Rows.InsertAt(emptyRow1, 0);
                    if (WaterLevel.Tables[0].Rows.Count > 0)
                    {
                        DataRow emptyRow2 = testsPerformed.Tables["Tests"].NewRow();
                        emptyRow2["TestName"] = "Water Level";
                        emptyRow2["StepNumber"] = "  ";
                        testsPerformed.Tables["Tests"].Rows.InsertAt(emptyRow2, 1);
                    }
                    testsPerformed.Tables["Tests"].Columns.Add("ForList", typeof(string), "StepNumber + ' - '+ TestName");

                    this.comboBox2.DisplayMember = "ForList";
                    this.comboBox2.ValueMember = "StepNumber";
                    this.comboBox2.DataSource = testsPerformed.Tables["Tests"];
                }
                catch(Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                

            }
        }


        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            // Clear the report type selection
            comboBox3.SelectedIndex = -1;
            // load data in the report specific methods below
            if (comboBox2.Text == "   - Water Level")
            {
                //show the water level report...
                waterLevelSummary();
            }
            
        }

        private void waterLevelSummary()
        {
            if (!(comboBox2.SelectedIndex < 1)) // make sure we have a selection to act on...
            {
                // we need a data set..
                DataSet WaterLevels = new DataSet();
                // Open database containing all the battery data....
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                string strAccessSelect = @"SELECT WLID,WorkOrderNumber,Cell1,Cell2,Cell3,Cell4,Cell5,Cell6,Cell7,Cell8,Cell9,Cell10,Cell11,Cell12,Cell13,Cell14,Cell15,Cell16,Cell17,Cell18,Cell19,Cell20,Cell21,Cell22,Cell23,Cell24,AVE FROM WaterLevel WHERE WorkOrderNumber='" + comboBox1.Text + @"' ORDER BY WLID ASC";

                //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
                OleDbConnection myAccessConn = null;
                // try to open the DB
                try
                {
                    myAccessConn = new OleDbConnection(strAccessConn);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                        myDataAdapter.Fill(WaterLevels, "WaterLevel");
                        myAccessConn.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {

                }


                // Now that we have the data in WaterLevels lets pass it over to the matching report
                /*************************Load reportSet into reportSet  ************************/

                // bind datatable to report viewer
                this.reportViewer.Reset();
                this.reportViewer.ProcessingMode = ProcessingMode.Local;

                this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WaterLevel.rdlc";
                this.reportViewer.LocalReport.DataSources.Clear();

                this.reportViewer.LocalReport.EnableExternalImages = true;
                ReportParameter parameter = new ReportParameter("Path", "file:////" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\rp_logo.jpg");
                this.reportViewer.LocalReport.SetParameters(parameter);

                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", WaterLevels.Tables[0]));

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

                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("DataSet2", MetaDT));
                this.reportViewer.RefreshReport();

                /*********************************************************/

                // finally enable the reportview
                reportViewer.Enabled = true;

            }
        }


        private void comboBox3_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                switch(comboBox3.SelectedIndex)
                {
                    case 0:
                        workOrderSummary();
                        break;
                    case 1:
                        batReport();
                        break;
                    case 2:
                        cellData();
                        break;
                    case 3:
                        testSummary();
                        break;
                    case 4:
                        workOrderLog();
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
            if (!(comboBox1.SelectedIndex < 1)) // make sure we have a selection to act on...
            {

                DataTable dtAll = new DataTable();
                // FIRST CLEAR THE OLD DATA SET!
                reportSet.Clear();
                // Open database containing all the battery data....
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                string strAccessSelect = @"SELECT * FROM ScanData WHERE BWO='" + comboBox1.Text + @"' ORDER BY DATE ASC";

                //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
                OleDbConnection myAccessConn = null;
                // try to open the DB
                try
                {
                    myAccessConn = new OleDbConnection(strAccessConn);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                //  now try to access it
                try
                {
                    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                    reportSet = new DataSet();
                    lock (Main_Form.dataBaseLock)
                    {
                        myAccessConn.Open();
                        myDataAdapter.Fill(reportSet, "ScanData");
                        myAccessConn.Close();
                    }

                    //now come up with a mergeded table...
                    dtAll = new DataTable();
                    dtAll = reportSet.Tables[0].Copy();
                    DataTable dtTemp = testsPerformed.Tables[0].Copy();
                    dtTemp.Rows.Remove(dtTemp.Rows[0]);

                    string[] columns = { "DateStarted", "DateCompleted", "Technician", "TestName", "Charger", "Notes" };

                    // add cols to the output table
                    foreach (string colname in columns)
                    {
                        dtAll.Columns.Add(colname);
                    }

                    // now insert data into the destination table
                    foreach (DataRow sourcerow in dtTemp.Rows)
                    {
                        //find the matching row in dtAll.Rows
                        foreach (DataRow destRow in dtAll.Rows)
                        {
                            if (sourcerow["StepNumber"].ToString() == destRow["Step"].ToString())
                            {
                                // we got a match...
                                foreach (string colname in columns)
                                {
                                    destRow[colname] = sourcerow[colname];
                                }
                            }
                        }

                    }

                    // now get rid of the repeats...
                    for (int i = 0; i < dtAll.Rows.Count - 1; i++)
                    {
                        if (dtAll.Rows[i]["TestName"].ToString().Contains("Cap") && (float.Parse(dtAll.Rows[i]["ETIME"].ToString())*24*60) > 50.5 && (float.Parse(dtAll.Rows[i]["ETIME"].ToString())*24*60) <= 51.5)
                        {
                            // skip the record if it is the 51 min of a cap test
                        }
                        else if ((int)dtAll.Rows[i][3] < (int)dtAll.Rows[i + 1][3])
                        {
                            dtAll.Rows.Remove(dtAll.Rows[i]);
                            i--;
                        }
                    }


                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {

                }


                // Now that we have the data in reportSet along with testsPerformed lets pass it over to the matching report
                /*************************Load reportSet into reportSet  ************************/

                // bind datatable to report viewer
                this.reportViewer.Reset();
                this.reportViewer.ProcessingMode = ProcessingMode.Local;
                switch (testsPerformed.Tables[0].Rows[1][15].ToString())
                {
                    case "1":
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_20_PN.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_20_NP.rdlc"; }
                        break;
                    case "2":
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_19_PN.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_19_NP.rdlc"; }
                        break;
                    case "3":
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_11_PN.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_11_NP.rdlc"; }
                        break;
                    case "4":
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_07_PN.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_07_NP.rdlc"; }
                        break;
                    case "10":
                        this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_04_BAT.rdlc";
                        break;
                    case "21":
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_21_PN.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_21_NP.rdlc"; }
                        break;
                    default:
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_24_PN.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_24_NP.rdlc"; }
                        break;
                }// end switch



                this.reportViewer.LocalReport.DataSources.Clear();
                this.reportViewer.LocalReport.EnableExternalImages = true;
                ReportParameter parameter = new ReportParameter("Path", "file:////" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\rp_logo.jpg");
                this.reportViewer.LocalReport.SetParameters(parameter);


                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("reportSet", reportSet.Tables[0]));
                this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("WOSumSet", dtAll));
                
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

        //private void workOrderSummary()
        //{
        //    reportSet = new  DataSet();

        //    if (!(comboBox1.SelectedIndex < 1)) // make sure we have a selection to act on...
        //    {
        //        // FIRST CLEAR THE OLD DATA SET!
        //        reportSet.Clear();
        //        // Open database containing all the battery data....
        //        string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
        //        string strAccessSelect = @"SELECT DATE FROM ScanData WHERE BWO='" + comboBox1.Text + "' ORDER BY RDG ASC";

        //        //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
        //        OleDbConnection myAccessConn = null;
        //        // try to open the DB
        //        try
        //        {
        //            myAccessConn = new OleDbConnection(strAccessConn);
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            return;
        //        }
        //        //  now try to access it
        //        try
        //        {
        //            OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
        //            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

        //            lock (Main_Form.dataBaseLock)
        //            {
        //                myAccessConn.Open();
        //                myDataAdapter.Fill(reportSet, "ScanData");
        //                myAccessConn.Close();
        //            }

        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            return;
        //        }
        //        finally
        //        {
                    
        //        }
        //    }

        //    DataTable testTable = new DataTable();


        //    testTable = testsPerformed.Tables[0].Copy();
        //    testTable.Rows[0].Delete();

        //    if (!(comboBox1.SelectedIndex < 0)) // make sure we have a selection to act on...
        //    {

        //        // We have the data in testsPerformed, so lets pass it over to the matching report

        //        // bind datatable to report viewer
        //        this.reportViewer.Reset();
        //        this.reportViewer.ProcessingMode = ProcessingMode.Local;
        //        this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WorkOrderSummary.rdlc";
        //        this.reportViewer.LocalReport.DataSources.Clear();

        //        this.reportViewer.LocalReport.EnableExternalImages = true;
        //        ReportParameter parameter = new ReportParameter("Path", "file:////" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\rp_logo.jpg");
        //        this.reportViewer.LocalReport.SetParameters(parameter);

        //        this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("reportSet", reportSet.Tables[0]));

        //        /*************************Load testsPerformed into Tests  ************************/


        //        this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("Tests", testTable));

        //        /*************************Load Global data into MetaData data Table ***********************
        //        // create datatable
        //        DataTable MetaDT = new DataTable("MetaData");
        //        MetaDT.Clear();

        //        // add columns
        //        MetaDT.Columns.Add("gBusinessName", typeof(string));
        //        MetaDT.Columns.Add("gUseF", typeof(string));
        //        MetaDT.Columns.Add("gPos2Neg", typeof(string));
        //        MetaDT.Columns.Add("testComboName", typeof(string));
        //        MetaDT.Columns.Add("cellsCable", typeof(string));
        //        MetaDT.Columns.Add("shuntCable", typeof(string));
        //        MetaDT.Columns.Add("tempCable", typeof(string));

        //        // insert data rows
        //        MetaDT.Rows.Add(GlobalVars.businessName, GlobalVars.useF, GlobalVars.Pos2Neg, comboBox2.Text, testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][15].ToString(), testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][16].ToString(), testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][17].ToString());
        //        */

        //        DataTable MetaDT = new DataTable("MetaData");
        //        this.reportViewer.LocalReport.DataSources.Add(new ReportDataSource("MetaData", MetaDT));

        //        this.reportViewer.RefreshReport();

        //        /*********************************************************/

        //        // finally enable the reportview
        //        reportViewer.Enabled = true;

        //    }

        //}

        private void workOrderLog()
        {
            reportSet = new DataSet();

            if (!(comboBox1.SelectedIndex < 1)) // make sure we have a selection to act on...
            {
                // FIRST CLEAR THE OLD DATA SET!
                reportSet.Clear();
                // Open database containing all the battery data....
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                string strAccessSelect = @"SELECT DATE FROM ScanData WHERE BWO='" + comboBox1.Text + "'  ORDER BY RDG ASC";

                //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
                OleDbConnection myAccessConn = null;
                // try to open the DB
                try
                {
                    myAccessConn = new OleDbConnection(strAccessConn);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                //  now try to access it
                try
                {
                    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                    reportSet = new DataSet();
                    lock (Main_Form.dataBaseLock)
                    {
                        myAccessConn.Open();
                        myDataAdapter.Fill(reportSet, "ScanData");
                        myAccessConn.Close();
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    
                }
            }

            DataTable testTable = new DataTable();

            
            testTable = testsPerformed.Tables[0].Copy();
            testTable.Rows[0].Delete();

            // also get rid of the water level test
            if(testTable.Rows[0][5].ToString().Contains("Water")){
                testTable.Rows[0].Delete();
            }

            if (!(comboBox1.SelectedIndex < 0)) // make sure we have a selection to act on...
            {
              
                // We have the data in testsPerformed, so lets pass it over to the matching report

                // bind datatable to report viewer
                this.reportViewer.Reset();

                this.reportViewer.ProcessingMode = ProcessingMode.Local;
                this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WorkOrderLog.rdlc"; 
                this.reportViewer.LocalReport.DataSources.Clear();

                this.reportViewer.LocalReport.EnableExternalImages = true;
                ReportParameter parameter = new ReportParameter("Path", "file:////" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\rp_logo.jpg");
                this.reportViewer.LocalReport.SetParameters(parameter);

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
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
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
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                //  now try to access it
                try
                {
                    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                    reportSet = new DataSet();
                    lock (Main_Form.dataBaseLock)
                    {
                        myAccessConn.Open();
                        myDataAdapter.Fill(reportSet, "ScanData");
                        myAccessConn.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    
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
                this.reportViewer.LocalReport.EnableExternalImages = true;
                ReportParameter parameter = new ReportParameter("Path", "file:////" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\rp_logo.jpg");
                this.reportViewer.LocalReport.SetParameters(parameter);



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
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                string strAccessSelect = @"SELECT DATE,RDG,ETIME,CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24 FROM ScanData WHERE BWO='" + comboBox1.Text + @"' AND STEP='" + comboBox2.Text.Substring(0, 2) + @"'  ORDER BY RDG ASC";

                //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
                OleDbConnection myAccessConn = null;
                // try to open the DB
                try
                {
                    myAccessConn = new OleDbConnection(strAccessConn);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                //  now try to access it
                try
                {
                    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                    reportSet = new DataSet();
                    lock (Main_Form.dataBaseLock)
                    {
                        myAccessConn.Open();
                        myDataAdapter.Fill(reportSet, "ScanData");
                        myAccessConn.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    
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
                    case "2":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN19.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP19.rdlc"; }
                        break;
                    case "3":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN11.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP11.rdlc"; }
                        break;
                    case "4":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN7.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP7.rdlc"; }
                        break;
                    case "21":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN21.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP21.rdlc"; }
                        break;
                    default:
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN24.rdlc"; }
                        else { this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP24.rdlc"; }
                        break;
                }// end switch

                this.reportViewer.LocalReport.DataSources.Clear();
                this.reportViewer.LocalReport.EnableExternalImages = true;
                ReportParameter parameter = new ReportParameter("Path", "file:////" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\rp_logo.jpg");
                this.reportViewer.LocalReport.SetParameters(parameter);



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
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                string strAccessSelect = @"SELECT RDG,DATE,ETIME,CUR1,VB1,VB2,VB3,VB4,BT1,BT2,BT3,BT4,REF,BT5 FROM ScanData WHERE BWO='" + comboBox1.Text + @"' AND STEP='" + comboBox2.Text.Substring(0, 2) + @"'  ORDER BY RDG ASC";

                //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
                OleDbConnection myAccessConn = null;
                // try to open the DB
                try
                {
                    myAccessConn = new OleDbConnection(strAccessConn);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                //  now try to access it
                try
                {
                    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                    reportSet = new DataSet();
                    lock (Main_Form.dataBaseLock)
                    {
                        myAccessConn.Open();
                        myDataAdapter.Fill(reportSet, "ScanData");
                        myAccessConn.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {

                }


                // Now that we have the data in reportSet along with testsPerformed lets pass it over to the matching report
                /*************************Load reportSet into reportSet  ************************/

                // bind datatable to report viewer
                this.reportViewer.Reset();
                this.reportViewer.ProcessingMode = ProcessingMode.Local;

                switch (testsPerformed.Tables[0].Rows[comboBox2.SelectedIndex][15].ToString())
                {
                    case "31":
                    case "10":
                        this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.BatteryData4.rdlc";
                        break;
                    case "4":
                        this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.BatteryData3.rdlc";
                        break;
                    case "3":
                        this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.BatteryData2.rdlc";
                        break;
                    default:
                        this.reportViewer.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.BatteryData1.rdlc";
                        break;
                }// end switch

                this.reportViewer.LocalReport.DataSources.Clear();
                this.reportViewer.LocalReport.EnableExternalImages = true;
                ReportParameter parameter = new ReportParameter("Path", "file:////" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\rp_logo.jpg");
                this.reportViewer.LocalReport.SetParameters(parameter);

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

        }

        private void Reports_Form_Shown(object sender, EventArgs e)
        {

        }


    }
}
