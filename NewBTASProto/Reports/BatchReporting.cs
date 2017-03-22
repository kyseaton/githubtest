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
using System.IO;
using Microsoft.Reporting.WinForms;
using System.Threading;

namespace NewBTASProto
{
    public partial class BatchReporting : Form
    {

        //form wide dataSets used to fill in the report
        DataSet reportSet = new DataSet();
        DataSet testsPerformed = new DataSet();
        string ComboText = "";
        bool cb1;
        bool cb2;
        bool cb3;
        bool cb4;
        bool cb5;
        bool cb6;
        bool cb7;   // Close work order
        bool cb8;   // Update Completion Date

        //for startup selection
        string curWorkOrder = "";

        public BatchReporting( string currentWorkOrder = "")
        {
            InitializeComponent();
            // we need to split up the work orders if we have multiple work orders on a single line...
            string tempWOS = currentWorkOrder;
            char[] delims = { ' ' };
            string[] A = tempWOS.Split(delims);
            curWorkOrder = A[0];

        }

        private void BatchReporting_Load(object sender, EventArgs e)
        {
            // Load all of the work orders into the combo box.
            loadWorkOrderLists();
            checkBox1.Checked = GlobalVars.cb1;
            checkBox2.Checked = GlobalVars.cb2;
            checkBox3.Checked = GlobalVars.cb3;
            checkBox4.Checked = GlobalVars.cb4;
            checkBox5.Checked = GlobalVars.cb5;
            checkBox6.Checked = GlobalVars.cb6;
            checkBox7.Checked = GlobalVars.cbComplete;
            checkBox8.Checked = GlobalVars.cbUpdateCompleteDate;

        }

        private void loadWorkOrderLists()
        {

            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT WorkOrderNumber FROM WorkOrders";



            DataSet workOrderList = new DataSet();
            OleDbConnection myAccessConn;
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
                    myDataAdapter.Fill(workOrderList, "ScanData");
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

            DataRow emptyRow1 = workOrderList.Tables["ScanData"].NewRow();
            emptyRow1["WorkOrderNumber"] = "";
            workOrderList.Tables["ScanData"].Rows.InsertAt(emptyRow1, 0);

            this.comboBox1.DisplayMember = "WorkOrderNumber";
            this.comboBox1.ValueMember = "WorkOrderNumber";
            this.comboBox1.DataSource = workOrderList.Tables["ScanData"];

        }

        //this is where we are going to save our data...
        string folder = "";

        private void button1_Click(object sender, EventArgs e)
        {
            // launch the folder picker to figure out where the reports are to be saved...
            if (comboBox1.Text == "")
            {
                MessageBox.Show(this, "Please Select a Work Order", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ComboText = comboBox1.Text;
            cb1 = checkBox1.Checked;
            cb2 = checkBox2.Checked;
            cb3 = checkBox3.Checked;
            cb4 = checkBox4.Checked;
            cb5 = checkBox5.Checked;
            cb6 = checkBox6.Checked;
            cb7 = checkBox7.Checked;
            cb8 = checkBox8.Checked;

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                folder = folderBrowserDialog1.SelectedPath;
                button1.Enabled = false;

                // do this on a thread pool thread...
                ThreadPool.QueueUserWorkItem(s =>
                {
                    try
                    {
                        //first we need to fill in tests performed...
                        testsPerformed = new DataSet();
                        // Open database containing all the battery data....
                        string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                        string strAccessSelect = @"SELECT * FROM Tests WHERE WorkOrderNumber='" + ComboText + @"' ORDER BY StepNumber ASC";

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
                            button1.Enabled = true;
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
                            this.Invoke((MethodInvoker)delegate()
                            {
                                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                button1.Enabled = true;
                            });
                            return;
                        }
                        finally
                        {

                        }

                        //Go through the subs and try to generate the report pdfs directly...
                        if (cb1)
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                label2.Text = "Generating Work Order Summary";
                            });
                            workOrderSummary();
                        }
                        if (cb2)
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                label2.Text = "Generating Water Level Report";
                            });
                            waterLevelSummary();
                        }
                        if (cb3)
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                label2.Text = "Generating Bat Data Reports";
                            });
                            batReport();
                        }
                        if (cb4)
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                label2.Text = "Generating Cell Data Reports";
                            });
                            cellData();
                        }
                        if (cb5)
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                label2.Text = "Generating Test Summary Reports";
                            });
                            testSummary();
                        }
                        if (cb6)
                        {
                            this.Invoke((MethodInvoker)delegate()
                            {
                                label2.Text = "Generating Work Order Log";
                            });
                            workOrderLog();
                        }
                        if (cb7)
                        {
                            //close the work order
                            this.Invoke((MethodInvoker)delegate()
                            {
                                label2.Text = "Closing Work Order";
                            });
                            closeWO(ComboText);
                            Thread.Sleep(500);
                        }
                        if (cb8)
                        {
                            //update the completion date for the work order
                            this.Invoke((MethodInvoker)delegate()
                            {
                                label2.Text = "Updating Completion Date";
                            });
                            updateCompleteDate(ComboText);
                            Thread.Sleep(500);
                        }
                        this.Invoke((MethodInvoker)delegate()
                        {
                            label2.Text = "";
                            button1.Enabled = true;
                            MessageBox.Show(this, "Reports Created In:  " + folder);
                        });
                    }
                    catch (Exception ex)
                    {
                        this.Invoke((MethodInvoker)delegate()
                        {
                            MessageBox.Show(this, "Error. Was not able to generate reports:  " + Environment.NewLine + ex.ToString());
                            button1.Enabled = true;
                        });
                    }
                });
            }
        }

        private void closeWO(string workOrderToClose)
        {
            // set up the db Connection
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            OleDbConnection conn = new OleDbConnection(connectionString);

            // Also update the model in the other tables!
            string cmdStr = "UPDATE WorkOrders SET OrderStatus='Closed' WHERE WorkOrderNumber='" + workOrderToClose + "' AND OrderStatus = 'Open'";
            OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
            lock (Main_Form.dataBaseLock)
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }
        }

        private void updateCompleteDate(string workOrderToUpdateDateCompleted)
        {
            // set up the db Connection
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            OleDbConnection conn = new OleDbConnection(connectionString);

            // Also update the model in the other tables!
            string cmdStr = "UPDATE WorkOrders SET DateCompleted='" + System.DateTime.Now.Date.ToString() + "' WHERE WorkOrderNumber='" + workOrderToUpdateDateCompleted + "'";
            OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
            lock (Main_Form.dataBaseLock)
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }
        }

        private void waterLevelSummary()
        {
            // create a temp report viewer...
            ReportViewer reportViewer1 = new ReportViewer();
            // we need a data set..
            DataSet WaterLevels = new DataSet();
            // Open database containing all the battery data....
            string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            string strAccessSelect = @"SELECT WLID,WorkOrderNumber,Cell1,Cell2,Cell3,Cell4,Cell5,Cell6,Cell7,Cell8,Cell9,Cell10,Cell11,Cell12,Cell13,Cell14,Cell15,Cell16,Cell17,Cell18,Cell19,Cell20,Cell21,Cell22,Cell23,Cell24,AVE FROM WaterLevel WHERE WorkOrderNumber='" + ComboText + @"' ORDER BY WLID ASC";

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
                    myDataAdapter.Fill(WaterLevels, "WaterLevel");
                    myAccessConn.Close();
                }
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate()
                {
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
                return;
            }
            finally
            {

            }

            if (WaterLevels.Tables[0].Rows.Count < 1)
            {
                //we don't have any water level data...
                return;
            }


            // Now that we have the data in WaterLevels lets pass it over to the matching report
            /*************************Load reportSet into reportSet  ************************/

            // bind datatable to report viewer
            reportViewer1.Reset();
            reportViewer1.ProcessingMode = ProcessingMode.Local;

            reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WaterLevel.rdlc";
            reportViewer1.LocalReport.DataSources.Clear();

            reportViewer1.LocalReport.EnableExternalImages = true;
            ReportParameter parameter = new ReportParameter("Path", "file:////" + GlobalVars.folderString + @"\BTAS16_DB\rp_logo.jpg");
            reportViewer1.LocalReport.SetParameters(parameter);

            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", WaterLevels.Tables[0]));

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
            MetaDT.Rows.Add(GlobalVars.businessName, GlobalVars.useF, GlobalVars.Pos2Neg, ComboText);

            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet2", MetaDT));

            //now we will write the report to file...
            Warning[] warnings;
            string[] streamids;
            string mimeType;
            string encoding;
            string filenameExtension;

            byte[] bytes = reportViewer1.LocalReport.Render("PDF", null, out mimeType, out encoding, out filenameExtension, out streamids, out warnings);
            using (FileStream fs = new FileStream(folder + "/" + ComboText + "_Water_Level.pdf", FileMode.Create))
            {
                fs.Write(bytes, 0, bytes.Length);
            }

            //reportViewer1.RefreshReport();

            /*********************************************************/

            // finally enable the reportview
            //reportViewer1.Enabled = true;

        }

        private void workOrderSummary()
        {

            // create a temp report viewer...
            ReportViewer reportViewer1 = new ReportViewer();

            DataTable dtAll = new DataTable();
            // FIRST CLEAR THE OLD DATA SET!
            reportSet.Clear();
            // Open database containing all the battery data....
            string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            string strAccessSelect = @"SELECT * FROM ScanData WHERE BWO='" + ComboText + @"' ORDER BY DATE ASC";

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

                reportSet = new DataSet();
                lock (Main_Form.dataBaseLock)
                {
                    myAccessConn.Open();
                    myDataAdapter.Fill(reportSet, "ScanData");
                    myAccessConn.Close();
                }

                //now come up with a merged table...
                dtAll = new DataTable();
                dtAll = reportSet.Tables[0].Copy();
                DataTable dtTemp = new DataTable(); 
                dtTemp = testsPerformed.Tables[0].Copy();

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
                    if (dtAll.Rows[i]["TestName"].ToString().Contains("Cap") && (GetDouble(dtAll.Rows[i]["ETIME"].ToString()) * 24 * 60) > 50.5 && (GetDouble(dtAll.Rows[i]["ETIME"].ToString()) * 24 * 60) <= 51.5)
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
                this.Invoke((MethodInvoker)delegate()
                {
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
                return;
            }
            finally
            {

            }

            //now lets get the water data
            // we need a data set..
            DataSet WaterLevels = new DataSet();
            // Open database containing all the battery data....
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT WLID,WorkOrderNumber,Cell1,Cell2,Cell3,Cell4,Cell5,Cell6,Cell7,Cell8,Cell9,Cell10,Cell11,Cell12,Cell13,Cell14,Cell15,Cell16,Cell17,Cell18,Cell19,Cell20,Cell21,Cell22,Cell23,Cell24,AVE FROM WaterLevel WHERE WorkOrderNumber='" + ComboText + @"' ORDER BY WLID ASC";

            //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
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
                    myDataAdapter.Fill(WaterLevels, "WaterLevel");
                    myAccessConn.Close();
                }
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate()
                {
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
                return;

            }
            finally
            {

            }

            //now lets get the battery serial number and part number
            // we need a data set..
            DataSet BatInfo = new DataSet();
            // Open database containing all the battery data....
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT BatteryModel,BatterySerialNumber FROM WorkOrders WHERE WorkOrderNumber='" + ComboText + @"'";

            //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
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
                    myDataAdapter.Fill(BatInfo, "BatInfo");
                    myAccessConn.Close();
                }
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate()
                {
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
                return;
            }
            finally
            {

            }

            // Now that we have the data in reportSet along with testsPerformed lets pass it over to the matching report
            /*************************Load reportSet into reportSet  ************************/

            // bind datatable to report viewer
            reportViewer1.Reset();
            reportViewer1.ProcessingMode = ProcessingMode.Local;
            switch (testsPerformed.Tables[0].Rows[0][15].ToString())
            {
                case "1":
                    if (GlobalVars.Pos2Neg == true) { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_20_PN.rdlc"; }
                    else { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_20_NP.rdlc"; }
                    break;
                case "2":
                    if (GlobalVars.Pos2Neg == true) { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_19_PN.rdlc"; }
                    else { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_19_NP.rdlc"; }
                    break;
                case "3":
                    if (GlobalVars.Pos2Neg == true) { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_11_PN.rdlc"; }
                    else { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_11_NP.rdlc"; }
                    break;
                case "4":
                    if (GlobalVars.Pos2Neg == true) { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_07_PN.rdlc"; }
                    else { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_07_NP.rdlc"; }
                    break;
                case "10":
                    reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_04_BAT.rdlc";
                    break;
                case "11":
                    reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_02_BAT.rdlc";
                    break;
                case "21":
                    if (GlobalVars.Pos2Neg == true) { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_21_PN.rdlc"; }
                    else { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_21_NP.rdlc"; }
                    break;
                case "22":
                    if (GlobalVars.Pos2Neg == true) { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_22_PN.rdlc"; }
                    else { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_22_NP.rdlc"; }
                    break;
                default:
                    if (GlobalVars.Pos2Neg == true) { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_24_PN.rdlc"; }
                    else { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WOTestSum_24_NP.rdlc"; }
                    break;
            }// end switch



            reportViewer1.LocalReport.DataSources.Clear();
            reportViewer1.LocalReport.EnableExternalImages = true;
            ReportParameter parameter = new ReportParameter("Path", "file:////" + GlobalVars.folderString + @"\BTAS16_DB\rp_logo.jpg");
            reportViewer1.LocalReport.SetParameters(parameter);


            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("reportSet", reportSet.Tables[0]));
            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("WOSumSet", dtAll));

            /*************************Load testsPerformed into Tests  ************************/

            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("Tests", testsPerformed.Tables[0]));

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
            MetaDT.Rows.Add(GlobalVars.businessName, GlobalVars.useF, GlobalVars.Pos2Neg, testsPerformed.Tables[0].Rows[0][4].ToString() + " - " + testsPerformed.Tables[0].Rows[0][5].ToString(), testsPerformed.Tables[0].Rows[0][15].ToString(), testsPerformed.Tables[0].Rows[0][16].ToString(), testsPerformed.Tables[0].Rows[0][17].ToString());

            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("MetaData", MetaDT));

            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("WaterLevelData", WaterLevels.Tables[0]));
            //now we will write the report to file...
            Warning[] warnings;
            string[] streamids;
            string mimeType;
            string encoding;
            string filenameExtension;

            byte[] bytes = reportViewer1.LocalReport.Render("PDF", null, out mimeType, out encoding, out filenameExtension, out streamids, out warnings);
            using (FileStream fs = new FileStream(folder + "/" + ComboText + "_WO_SUM.pdf", FileMode.Create))
            {
                fs.Write(bytes, 0, bytes.Length);
            }

            //reportViewer1.RefreshReport();

            /*********************************************************/

            // finally enable the reportview
            //reportViewer1.Enabled = true;

        }

        private void batReport()
        {

            // create a temp report viewer...
            ReportViewer reportViewer1 = new ReportViewer();

            for (int j = 0; j < testsPerformed.Tables[0].Rows.Count; j++)
            {
                // FIRST CLEAR THE OLD DATA SET!
                reportSet.Clear();
                // Open database containing all the battery data....
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                string strAccessSelect = @"SELECT RDG,DATE,ETIME,CUR1,VB1,VB2,VB3,VB4,BT1,BT2,BT3,BT4,REF,BT5 FROM ScanData WHERE BWO='" + testsPerformed.Tables[0].Rows[j][2].ToString() + @"' AND STEP='" + testsPerformed.Tables[0].Rows[j][4].ToString() + @"'  ORDER BY RDG ASC";

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
                    this.Invoke((MethodInvoker)delegate()
                    {
                        MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                    return;
                }
                finally
                {

                }


                // Now that we have the data in reportSet along with testsPerformed lets pass it over to the matching report
                /*************************Load reportSet into reportSet  ************************/

                // bind datatable to report viewer
                reportViewer1.Reset();
                reportViewer1.ProcessingMode = ProcessingMode.Local;

                switch (testsPerformed.Tables[0].Rows[j][15].ToString())
                {
                    case "31":
                    case "10":
                        reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.BatteryData4.rdlc";
                        break;
                    case "4":
                        reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.BatteryData3.rdlc";
                        break;
                    case "3":
                        reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.BatteryData2.rdlc";
                        break;
                    default:
                        reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.BatteryData1.rdlc";
                        break;
                }// end switch

                reportViewer1.LocalReport.DataSources.Clear();
                reportViewer1.LocalReport.EnableExternalImages = true;
                ReportParameter parameter = new ReportParameter("Path", "file:////" + GlobalVars.folderString + @"\BTAS16_DB\rp_logo.jpg");
                reportViewer1.LocalReport.SetParameters(parameter);

                reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("reportSet", reportSet.Tables[0]));

                /*************************Load testsPerformed into Tests  ************************/

                reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("Tests", testsPerformed.Tables[0]));

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
                MetaDT.Rows.Add(GlobalVars.businessName, GlobalVars.useF, GlobalVars.Pos2Neg, testsPerformed.Tables[0].Rows[j][4].ToString() + " - " + testsPerformed.Tables[0].Rows[j][5].ToString(), testsPerformed.Tables[0].Rows[j][15].ToString(), testsPerformed.Tables[0].Rows[j][16].ToString(), testsPerformed.Tables[0].Rows[j][17].ToString());

                reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("MetaData", MetaDT));

                //now we will write the report to file...
                Warning[] warnings;
                string[] streamids;
                string mimeType;
                string encoding;
                string filenameExtension;

                byte[] bytes = reportViewer1.LocalReport.Render("PDF", null, out mimeType, out encoding, out filenameExtension, out streamids, out warnings);
                using (FileStream fs = new FileStream(folder + "/" + ComboText + "_" + testsPerformed.Tables[0].Rows[j][4].ToString() + "_BAT_DATA.pdf", FileMode.Create))
                {
                    fs.Write(bytes, 0, bytes.Length);
                }

                //reportViewer1.RefreshReport();

                /*********************************************************/

                // finally enable the reportview
                //reportViewer1.Enabled = true;

            }

        }

        private void cellData()
        {

            // create a temp report viewer...
            ReportViewer reportViewer1 = new ReportViewer();

            for (int j = 0; j < testsPerformed.Tables[0].Rows.Count; j++)
            {
                // FIRST CLEAR THE OLD DATA SET!
                reportSet.Clear();
                // Open database containing all the battery data....
                string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                string strAccessSelect = @"SELECT DATE,RDG,ETIME,CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24 FROM ScanData WHERE BWO='" + testsPerformed.Tables[0].Rows[j][2].ToString() + @"' AND STEP='" + testsPerformed.Tables[0].Rows[j][4].ToString() + @"'  ORDER BY RDG ASC";

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
                    this.Invoke((MethodInvoker)delegate()
                    {
                        MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                    return;
                }
                finally
                {

                }


                // Now that we have the data in reportSet along with testsPerformed lets pass it over to the matching report
                /*************************Load reportSet into reportSet  ************************/

                // bind datatable to report viewer
                reportViewer1.Reset();
                reportViewer1.ProcessingMode = ProcessingMode.Local;
                switch (testsPerformed.Tables[0].Rows[j][15].ToString())
                {
                    case "1":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN20.rdlc"; }
                        else { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP20.rdlc"; }
                        break;
                    case "2":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN19.rdlc"; }
                        else { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP19.rdlc"; }
                        break;
                    case "3":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN11.rdlc"; }
                        else { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP11.rdlc"; }
                        break;
                    case "4":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN7.rdlc"; }
                        else { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP7.rdlc"; }
                        break;
                    case "21":
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN21.rdlc"; }
                        else { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP21.rdlc"; }
                        break;
                    default:
                        // update the cells value
                        if (GlobalVars.Pos2Neg == true) { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataPN24.rdlc"; }
                        else { reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.CellDataNP24.rdlc"; }
                        break;
                }// end switch

                reportViewer1.LocalReport.DataSources.Clear();
                reportViewer1.LocalReport.EnableExternalImages = true;
                ReportParameter parameter = new ReportParameter("Path", "file:////" + GlobalVars.folderString + @"\BTAS16_DB\rp_logo.jpg");
                reportViewer1.LocalReport.SetParameters(parameter);



                reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("reportSet", reportSet.Tables[0]));

                /*************************Load testsPerformed into Tests  ************************/

                reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("Tests", testsPerformed.Tables[0]));

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
                MetaDT.Rows.Add(GlobalVars.businessName, GlobalVars.useF, GlobalVars.Pos2Neg, testsPerformed.Tables[0].Rows[j][4].ToString() + " - " + testsPerformed.Tables[0].Rows[j][5].ToString(), testsPerformed.Tables[0].Rows[j][15].ToString(), testsPerformed.Tables[0].Rows[j][16].ToString(), testsPerformed.Tables[0].Rows[j][17].ToString());

                reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("MetaData", MetaDT));


                //now we will write the report to file...
                Warning[] warnings;
                string[] streamids;
                string mimeType;
                string encoding;
                string filenameExtension;

                byte[] bytes = reportViewer1.LocalReport.Render("PDF", null, out mimeType, out encoding, out filenameExtension, out streamids, out warnings);
                using (FileStream fs = new FileStream(folder + "/" + ComboText + "_" + testsPerformed.Tables[0].Rows[j][4].ToString() + "_CELL_DATA.pdf", FileMode.Create))
                {
                    fs.Write(bytes, 0, bytes.Length);
                }

                
                //reportViewer1.RefreshReport();

                /*********************************************************/

                // finally enable the reportview
                //reportViewer1.Enabled = true;

            }
        }

        private void testSummary()
        {
            try
            {

                // create a temp report viewer...
                ReportViewer reportViewer1 = new ReportViewer();
                reportSet = new DataSet();

                for (int j = 0; j < testsPerformed.Tables[0].Rows.Count; j++)
                {
                    // FIRST CLEAR THE OLD DATA SET!
                    reportSet.Clear();
                    // Open database containing all the battery data....
                    string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                    string strAccessSelect = @"SELECT DATE,RDG,ETIME,CEL01,CEL02,CEL03,CEL04,CEL05,CEL06,CEL07,CEL08,CEL09,CEL10,CEL11,CEL12,CEL13,CEL14,CEL15,CEL16,CEL17,CEL18,CEL19,CEL20,CEL21,CEL22,CEL23,CEL24 FROM ScanData WHERE BWO='" + testsPerformed.Tables[0].Rows[j][2].ToString() + @"' AND STEP='" + testsPerformed.Tables[0].Rows[j][4].ToString() + @"'";

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
                    //Now lets go through the data and come up with a pass fail for each cell
                    DataSet passFailSet = new DataSet();
                    passFailSet.Tables.Add("PassFail");
                    passFailSet.Tables[0].Columns.Add("Station");
                    passFailSet.Tables[0].Columns.Add("ETIME");
                    passFailSet.Tables[0].Columns.Add("RDG");
                    passFailSet.Tables[0].Columns.Add("CEL01");
                    passFailSet.Tables[0].Columns.Add("Ref");
                    passFailSet.Tables[0].Columns.Add("DATE");

                    bool charge = true;        // default to discharge

                    //how many cells do we need to display?
                    int cells = 24;
                    try
                    {
                        switch (int.Parse(testsPerformed.Tables[0].Rows[j][15].ToString()))
                        {
                            case 1:
                                cells = 20;
                                break;
                            case 2:
                                cells = 19;
                                break;
                            case 10:
                                cells = 0;
                                break;
                            case 31:
                                cells = 24;
                                break;
                            case 3:
                                cells = 11;
                                break;
                            case 4:
                                cells = 7;
                                break;
                            case 21:
                                cells = 21;
                                break;
                            case 22:
                                cells = 22;
                                break;
                            case 9:
                            case 11:
                                cells = 0;
                                break;
                            case 23:
                                cells = 20;
                                break;
                        }
                    }
                    catch
                    {
                        cells = 0;
                    }

                    if (testsPerformed.Tables[0].Rows[j][5].ToString().Contains("Cap") || testsPerformed.Tables[0].Rows[j][5].ToString().Contains("Dis") || testsPerformed.Tables[0].Rows[j][5].ToString() == "As Received")
                    { charge = false; }


                    if (GlobalVars.Pos2Neg == false)
                    {
                        if (charge)
                        {


                            for (int i = 0; i < cells; i++)
                            {
                                for (int q = 0; q < reportSet.Tables[0].Rows.Count; q++)
                                {
                                    //check to see where we get below 1. When we do log it in passFailSet
                                    if (float.Parse(reportSet.Tables[0].Rows[q][i + 3].ToString()) > 1.75)
                                    {
                                        passFailSet.Tables[0].Rows.Add(); // add a row
                                        passFailSet.Tables[0].Rows[i][0] = "Cell" + (i + 1).ToString();// record the cell number
                                        passFailSet.Tables[0].Rows[i][1] = reportSet.Tables[0].Rows[q][2];// record the time
                                        passFailSet.Tables[0].Rows[i][2] = reportSet.Tables[0].Rows[q][1];// record the rdg
                                        passFailSet.Tables[0].Rows[i][3] = reportSet.Tables[0].Rows[q][i + 3];// record the cell voltage
                                        passFailSet.Tables[0].Rows[i][4] = "FAIL! Overvoltage!";// record the status
                                        passFailSet.Tables[0].Rows[i][5] = reportSet.Tables[0].Rows[q][0];// record the time
                                        break; // move to the next cell
                                    }
                                }

                                //check to see if it didn't over volt passed everything
                                if (float.Parse(reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][i + 3].ToString()) <= 1.75)
                                {
                                    passFailSet.Tables[0].Rows.Add(); // add a row
                                    passFailSet.Tables[0].Rows[i][0] = "Cell" + (i + 1).ToString();// record the cell number
                                    passFailSet.Tables[0].Rows[i][1] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][2];// record the time
                                    passFailSet.Tables[0].Rows[i][2] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][1];// record the rdg
                                    passFailSet.Tables[0].Rows[i][3] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][i + 3];// record the cell voltage
                                    passFailSet.Tables[0].Rows[i][5] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][0];// record the time
                                    //check for under voltage
                                    if (float.Parse(reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][i + 3].ToString()) < 1)
                                    {
                                        passFailSet.Tables[0].Rows[i][4] = "FAIL!";// record the status
                                    }
                                    else
                                    {
                                        passFailSet.Tables[0].Rows[i][4] = "OK";// record the status
                                    }
                                }

                            }
                        }
                        else
                        {
                            for (int i = 0; i < cells; i++)
                            {
                                for (int q = 0; q < reportSet.Tables[0].Rows.Count; q++)
                                {
                                    //check to see where we get below 1. When we do log it in passFailSet
                                    if (float.Parse(reportSet.Tables[0].Rows[q][i + 3].ToString()) < 1)
                                    {
                                        passFailSet.Tables[0].Rows.Add(); // add a row
                                        passFailSet.Tables[0].Rows[i][0] = "Cell" + (i + 1).ToString();// record the cell number
                                        passFailSet.Tables[0].Rows[i][1] = reportSet.Tables[0].Rows[q][2];// record the time
                                        passFailSet.Tables[0].Rows[i][2] = reportSet.Tables[0].Rows[q][1];// record the rdg
                                        passFailSet.Tables[0].Rows[i][3] = reportSet.Tables[0].Rows[q][i + 3];// record the cell voltage
                                        passFailSet.Tables[0].Rows[i][4] = "FAIL!";// record the status
                                        passFailSet.Tables[0].Rows[i][5] = reportSet.Tables[0].Rows[q][0];// record the time
                                        break; // move to the next cell
                                    }
                                }

                                //check to see if it passed everything
                                if (float.Parse(reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][i + 3].ToString()) >= 1)
                                {
                                    passFailSet.Tables[0].Rows.Add(); // add a row
                                    passFailSet.Tables[0].Rows[i][0] = "Cell" + (i + 1).ToString();// record the cell number
                                    passFailSet.Tables[0].Rows[i][1] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][2];// record the time
                                    passFailSet.Tables[0].Rows[i][2] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][1];// record the rdg
                                    passFailSet.Tables[0].Rows[i][3] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][i + 3];// record the cell voltage
                                    passFailSet.Tables[0].Rows[i][4] = "OK";// record the status
                                    passFailSet.Tables[0].Rows[i][5] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][0];// record the time
                                }
                            }
                        }
                    }
                    else
                    {
                        if (charge)
                        {


                            for (int i = cells - 1; i >= 0; i--)
                            {
                                for (int q = 0; q < reportSet.Tables[0].Rows.Count; q++)
                                {
                                    //check to see where we get below 1. When we do log it in passFailSet
                                    if (float.Parse(reportSet.Tables[0].Rows[q][i + 3].ToString()) > 1.75)
                                    {
                                        passFailSet.Tables[0].Rows.Add(); // add a row
                                        passFailSet.Tables[0].Rows[cells - i - 1][0] = "Cell" + (cells - i).ToString();// record the cell number
                                        passFailSet.Tables[0].Rows[cells - i - 1][1] = reportSet.Tables[0].Rows[q][2];// record the time
                                        passFailSet.Tables[0].Rows[cells - i - 1][2] = reportSet.Tables[0].Rows[q][1];// record the rdg
                                        passFailSet.Tables[0].Rows[cells - i - 1][3] = reportSet.Tables[0].Rows[q][i + 3];// record the cell voltage
                                        passFailSet.Tables[0].Rows[cells - i - 1][4] = "FAIL! Overvoltage!";// record the status
                                        passFailSet.Tables[0].Rows[cells - i - 1][5] = reportSet.Tables[0].Rows[j][0];// record the time
                                        break; // move to the next cell
                                    }
                                }

                                //check to see if it didn't over volt passed everything
                                if (float.Parse(reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][i + 3].ToString()) <= 1.75)
                                {
                                    passFailSet.Tables[0].Rows.Add(); // add a row
                                    passFailSet.Tables[0].Rows[cells - i - 1][0] = "Cell" + (cells - i).ToString();// record the cell number
                                    passFailSet.Tables[0].Rows[cells - i - 1][1] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][2];// record the time
                                    passFailSet.Tables[0].Rows[cells - i - 1][2] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][1];// record the rdg
                                    passFailSet.Tables[0].Rows[cells - i - 1][3] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][i + 3];// record the cell voltage
                                    passFailSet.Tables[0].Rows[cells - i - 1][5] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][0];// record the time
                                    //check for under voltage
                                    if (float.Parse(reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][cells - i - 1 + 3].ToString()) < 1)
                                    {
                                        passFailSet.Tables[0].Rows[cells - i - 1][4] = "FAIL!";// record the status
                                    }
                                    else
                                    {
                                        passFailSet.Tables[0].Rows[cells - i - 1][4] = "OK";// record the status
                                    }
                                }

                            }
                        }
                        else
                        {
                            for (int i = cells - 1; i > -1; i--)
                            {
                                for (int q = 0; q < reportSet.Tables[0].Rows.Count; q++)
                                {
                                    //check to see where we get below 1. When we do log it in passFailSet
                                    if (float.Parse(reportSet.Tables[0].Rows[q][cells - i - 1 + 3].ToString()) < 1)
                                    {
                                        passFailSet.Tables[0].Rows.Add(); // add a row
                                        passFailSet.Tables[0].Rows[cells - i - 1][0] = "Cell" + (cells - i).ToString();// record the cell number
                                        passFailSet.Tables[0].Rows[cells - i - 1][1] = reportSet.Tables[0].Rows[q][2];// record the time
                                        passFailSet.Tables[0].Rows[cells - i - 1][2] = reportSet.Tables[0].Rows[q][1];// record the rdg
                                        passFailSet.Tables[0].Rows[cells - i - 1][3] = reportSet.Tables[0].Rows[q][i + 3];// record the cell voltage
                                        passFailSet.Tables[0].Rows[cells - i - 1][4] = "FAIL!";// record the status
                                        passFailSet.Tables[0].Rows[cells - i - 1][5] = reportSet.Tables[0].Rows[q][0];// record the time
                                        break; // move to the next cell
                                    }
                                }

                                //check to see if it passed everything
                                if (float.Parse(reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][cells - i - 1 + 3].ToString()) >= 1)
                                {
                                    passFailSet.Tables[0].Rows.Add(); // add a row
                                    passFailSet.Tables[0].Rows[cells - i - 1][0] = "Cell" + (cells - i).ToString();// record the cell number
                                    passFailSet.Tables[0].Rows[cells - i - 1][1] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][2];// record the time
                                    passFailSet.Tables[0].Rows[cells - i - 1][2] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][1];// record the rdg
                                    passFailSet.Tables[0].Rows[cells - i - 1][3] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][i + 3];// record the cell voltage
                                    passFailSet.Tables[0].Rows[cells - i - 1][4] = "OK";// record the status
                                    passFailSet.Tables[0].Rows[cells - i - 1][5] = reportSet.Tables[0].Rows[reportSet.Tables[0].Rows.Count - 1][0];// record the time
                                }
                            }
                        }
                    }



                    // Now that we have the data in reportSet along with testsPerformed lets pass it over to the matching report
                    /*************************Load reportSet into reportSet  ************************/

                    // bind datatable to report viewer
                    reportViewer1.Reset();
                    reportViewer1.ProcessingMode = ProcessingMode.Local;

                    reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.TestSummary.rdlc";


                    reportViewer1.LocalReport.DataSources.Clear();
                    reportViewer1.LocalReport.EnableExternalImages = true;
                    ReportParameter parameter = new ReportParameter("Path", "file:////" + GlobalVars.folderString + @"\BTAS16_DB\rp_logo.jpg");
                    reportViewer1.LocalReport.SetParameters(parameter);



                    reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("reportSet", passFailSet.Tables[0]));

                    /*************************Load testsPerformed into Tests  ************************/

                    reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("Tests", testsPerformed.Tables[0]));

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
                    MetaDT.Rows.Add(GlobalVars.businessName, GlobalVars.useF, GlobalVars.Pos2Neg, testsPerformed.Tables[0].Rows[j][4].ToString() + " - " + testsPerformed.Tables[0].Rows[j][5].ToString(), testsPerformed.Tables[0].Rows[j][15].ToString(), testsPerformed.Tables[0].Rows[j][16].ToString(), testsPerformed.Tables[0].Rows[j][17].ToString());

                    reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("MetaData", MetaDT));

                    //now we will write the report to file...
                    Warning[] warnings;
                    string[] streamids;
                    string mimeType;
                    string encoding;
                    string filenameExtension;

                    byte[] bytes = reportViewer1.LocalReport.Render("PDF", null, out mimeType, out encoding, out filenameExtension, out streamids, out warnings);
                    using (FileStream fs = new FileStream(folder + "/" + ComboText + "_" + testsPerformed.Tables[0].Rows[j][4].ToString() + "_TEST_SUM.pdf", FileMode.Create))
                    {
                        fs.Write(bytes, 0, bytes.Length);
                    }

                    //reportViewer1.RefreshReport();

                    /*********************************************************/

                    // finally enable the reportview
                    //reportViewer1.Enabled = true;

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void workOrderLog()
        {

            // create a temp report viewer...
            ReportViewer reportViewer1 = new ReportViewer();
            reportSet = new DataSet();

            // FIRST CLEAR THE OLD DATA SET!
            reportSet.Clear();
            // Open database containing all the battery data....
            string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
            string strAccessSelect = @"SELECT DATE FROM ScanData WHERE BWO='" + ComboText + "'  ORDER BY RDG ASC";

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
                this.Invoke((MethodInvoker)delegate()
                {
                    MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
                return;
            }
            finally
            {

            }


            DataTable testTable = new DataTable();


            testTable = testsPerformed.Tables[0].Copy();
            //testTable.Rows[0].Delete();


            // We have the data in testsPerformed, so lets pass it over to the matching report

            // bind datatable to report viewer
            reportViewer1.Reset();

            reportViewer1.ProcessingMode = ProcessingMode.Local;
            reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WorkOrderLog.rdlc";
            reportViewer1.LocalReport.DataSources.Clear();

            reportViewer1.LocalReport.EnableExternalImages = true;
            ReportParameter parameter = new ReportParameter("Path", "file:////" + GlobalVars.folderString + @"\BTAS16_DB\rp_logo.jpg");
            reportViewer1.LocalReport.SetParameters(parameter);

            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("reportSet", reportSet.Tables[0]));

            /*************************Load testsPerformed into Tests  ************************/


            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("Tests", testTable));

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
            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("MetaData", MetaDT));


            //now we will write the report to file...
            Warning[] warnings;
            string[] streamids;
            string mimeType;
            string encoding;
            string filenameExtension;

            byte[] bytes = reportViewer1.LocalReport.Render("PDF", null, out mimeType, out encoding, out filenameExtension, out streamids, out warnings);
            using (FileStream fs = new FileStream(folder + "/" + ComboText + "_Work_Order_LOG.pdf", FileMode.Create))
            {
                fs.Write(bytes, 0, bytes.Length);
            }


            //reportViewer1.RefreshReport();

            /*********************************************************/

            // finally enable the reportview

        }

        private void BatchReporting_FormClosing(object sender, FormClosingEventArgs e)
        {
            GlobalVars.cb1 = checkBox1.Checked;
            GlobalVars.cb2 = checkBox2.Checked;
            GlobalVars.cb3 = checkBox3.Checked;
            GlobalVars.cb4 = checkBox4.Checked;
            GlobalVars.cb5 = checkBox5.Checked;
            GlobalVars.cb6 = checkBox6.Checked;
            GlobalVars.cbComplete = checkBox7.Checked;
            GlobalVars.cbUpdateCompleteDate = checkBox8.Checked;
        }

        private void BatchReporting_Shown(object sender, EventArgs e)
        {
            comboBox1.Text = curWorkOrder;
            button1.Focus();
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

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {

        }
        
    }
}
