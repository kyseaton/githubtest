using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Threading;
using Microsoft.Reporting.WinForms;

namespace NewBTASProto
{

    public partial class WorkOrderReps : Form
    {
        // class wide variables
        DataSet WorkOrders = new DataSet();
        bool startup = true;


        public WorkOrderReps()
        {
            InitializeComponent();
            LoadData();
        }

        private void LoadData()
        {
            #region setup the binding

            // THIS SHOULD BE DONE ON A HELPER THREAD!!!

            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            if (toolStripCBWorkOrderStatus.Text != "" || toolStripCBCustomers.Text != "" || toolStripCBSerialNums.Text != "")
            {
                strAccessSelect = @"SELECT * FROM WorkOrders WHERE " +
                    (toolStripCBWorkOrderStatus.Text != "" ? ("OrderStatus='" + toolStripCBWorkOrderStatus.Text + "' ") : "") +
                    ((toolStripCBWorkOrderStatus.Text != "" && toolStripCBCustomers.Text != "") ? " AND " : "") +
                    (toolStripCBCustomers.Text != "" ? ("CustomerName='" + toolStripCBCustomers.Text + "' ") : "") +
                    ((toolStripCBCustomers.Text != "" && toolStripCBSerialNums.Text != "") ? " AND " : "") +
                    ((toolStripCBWorkOrderStatus.Text != "" && toolStripCBCustomers.Text == "" && toolStripCBSerialNums.Text != "") ? " AND " : "") +
                    (toolStripCBSerialNums.Text != "" ? ("BatterySerialNumber='" + toolStripCBSerialNums.Text + "' ") : "")+
                    "AND DateReceived BETWEEN #" + dateTimePicker1.Value.ToString() +@"# AND #" + dateTimePicker2.Value.ToString() + @"# ORDER BY WorkOrderNumber ASC";
            }
            else
            {
                strAccessSelect = @"SELECT * FROM WorkOrders WHERE DateReceived BETWEEN #" + dateTimePicker1.Value.ToString() +@"# AND #" + dateTimePicker2.Value.ToString() + @"# ORDER BY WorkOrderNumber ASC";
            }

            WorkOrders.Clear();
            OleDbConnection myAccessConn = null;
            // try to open the DB
            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to create a database connection. \n" + ex.Message);
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
                    myDataAdapter.Fill(WorkOrders);
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                return;
            }
            finally
            {
                
            }

            // Set the DataSource to the DataSet, and the DataMember

            #endregion
            if (startup)
            {
                #region setup the combo boxes


                //  Setup the drop down to contain all customers availible in the customer table

                // Open database containing all the customer names data....

                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                strAccessSelect = @"SELECT CustomerName FROM CUSTOMERS ORDER BY CustomerName ASC";

                DataSet Custs = new DataSet();
                myAccessConn = null;
                // try to open the DB
                try
                {
                    myAccessConn = new OleDbConnection(strAccessConn);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to create a database connection. \n" + ex.Message);
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
                        myDataAdapter.Fill(Custs);
                        myAccessConn.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    return;
                }
                finally
                {

                }

                List<string> Customers = Custs.Tables[0].AsEnumerable().Select(x => x[0].ToString()).Distinct().ToList();
                Customers.Sort();
                Customers.Insert(0, "");
                toolStripCBCustomers.DataSource = Customers;

                //           foreach (string x in Customers)
                //           {
                //               comboBox1.Items.Add(x);
                //           }

                //Now we'll set up the Battery Serial Number drop down, so the customer can re assign the battery associated with the work order

                // Open database containing all the customer names data....

                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                strAccessSelect = @"SELECT BatterySerialNumber FROM Batteries ORDER BY BatterySerialNumber ASC";

                DataSet Serials = new DataSet();
                myAccessConn = null;
                // try to open the DB
                try
                {
                    myAccessConn = new OleDbConnection(strAccessConn);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to create a database connection. \n" + ex.Message);
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
                        myDataAdapter.Fill(Serials);
                        myAccessConn.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(new Form() { TopMost = true }, "Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                    return;
                }
                finally
                {

                }

                List<string> SerialNums = Serials.Tables[0].AsEnumerable().Select(x => x[0].ToString()).Distinct().ToList();
                SerialNums.Sort();
                SerialNums.Insert(0, "");
                toolStripCBSerialNums.DataSource = SerialNums;

                #endregion
            }

            #region Now set up the master report..

            // bind datatable to report viewer
            this.reportViewer1.Reset();
            this.reportViewer1.ProcessingMode = ProcessingMode.Local;

            this.reportViewer1.LocalReport.EnableHyperlinks = true;
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.WorkOrderSum.rdlc";
            this.reportViewer1.LocalReport.DataSources.Clear();
            
            this.reportViewer1.LocalReport.EnableExternalImages = true;
            ReportParameter parameter = new ReportParameter("Path", "file:////" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\rp_logo.jpg");
            this.reportViewer1.LocalReport.SetParameters(parameter);

            this.reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("WorkOrders", WorkOrders.Tables[0]));

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
            MetaDT.Rows.Add(GlobalVars.businessName, "  ", " ", " ", " ", " ", " ");

            this.reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("MetaData", MetaDT));
            this.reportViewer1.RefreshReport();

            /*********************************************************/

            // finally enable the reportview
            reportViewer1.Enabled = true;
            #endregion

        }

        private void bindingSource1_DataError(object sender, BindingManagerDataErrorEventArgs e)
        {
            //here!
        }

        private void bindingSource1_DataMemberChanged(object sender, EventArgs e)
        {

        }

        private void bindingSource1_AddingNew(object sender, AddingNewEventArgs e)
        {

        }

        private void bindingSource1_BindingComplete(object sender, BindingCompleteEventArgs e)
        {

        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void bindingSource1_CurrentItemChanged(object sender, EventArgs e)
        {

        }

        private void bindingSource1_DataSourceChanged(object sender, EventArgs e)
        {

        }

        private void bindingSource1_ListChanged(object sender, ListChangedEventArgs e)
        {

        }

        private void bindingSource1_PositionChanged(object sender, EventArgs e)
        {
            //lastValid = false;
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void toolStripCBWorkOrders_SelectedIndexChanged(object sender, EventArgs e)
        {


        }


        private void UpdateReport()
        {
            

        }


        private void toolStripCBWorkOrderStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void toolStripCBCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void toolStripCBSerialNums_SelectedIndexChanged(object sender, EventArgs e)
        {
        }



        private void frmVEWorkOrders_Load(object sender, EventArgs e)
        {

            this.reportViewer1.RefreshReport();
        }

        private void bindingNavigatorMoveNextItem_Click(object sender, EventArgs e)
        {

        }


        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorMovePreviousItem_Click(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorMoveFirstItem_Click(object sender, EventArgs e)
        {

        }

        private void toolStripCBWorkOrders_TextChanged(object sender, EventArgs e)
        {

        }


        private void bindingNavigator1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void frmVEWorkOrders_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void bindingNavigator1_Validating(object sender, CancelEventArgs e)
        {

        }

        private void toolStripCBWorkOrderStatus_Enter(object sender, EventArgs e)
        {

        }

        private void toolStripCBWorkOrderStatus_Leave(object sender, EventArgs e)
        {

        }

        private void toolStripCBCustomers_Enter(object sender, EventArgs e)
        {
        }

        private void toolStripCBCustomers_Leave(object sender, EventArgs e)
        {

        }

        private void toolStripCBSerialNums_Enter(object sender, EventArgs e)
        {
        
        }

        private void toolStripCBSerialNums_Leave(object sender, EventArgs e)
        {

        }

        private void toolStripCBWorkOrders_Enter(object sender, EventArgs e)
        {

        }

        private void toolStripCBWorkOrders_Leave(object sender, EventArgs e)
        {

        }

        private void bindingNavigator1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void toolStripCBSerialNums_TextChanged(object sender, EventArgs e)
        {

        }

        private void toolStripCBCustomers_TextChanged(object sender, EventArgs e)
        {

        }

        private void toolStripCBWorkOrderStatus_TextChanged(object sender, EventArgs e)
        {

        }

        private void frmVEWorkOrders_Shown(object sender, EventArgs e)
        {
            startup = false;

        }

        private void comboBox1_Enter(object sender, EventArgs e)
        {

        }

        private void toolStripCBWorkOrderStatus_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            LoadData();
        }

        private void toolStripCBCustomers_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            LoadData();
        }

        private void toolStripCBSerialNums_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            LoadData();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            LoadData();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            LoadData();
        }

        private void reportViewer1_Hyperlink(object sender, HyperlinkEventArgs e)
        {
            string wo = e.Hyperlink;
            wo = wo.Substring(7);
            wo = wo.Substring(0, wo.Length - 1);

            Reports_Form rf = new Reports_Form("1", wo);
            rf.Owner = this.Owner;
            rf.Show();
            e.Cancel = true;
        }

    }
}
