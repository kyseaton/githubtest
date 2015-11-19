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
using System.Data.SqlClient;
using Microsoft.Reporting.WinForms;

namespace NewBTASProto
{
    public partial class Battery_Reports : Form
    {

        DataSet Bats = new DataSet();
       

        public Battery_Reports()
        {
            InitializeComponent();
            LoadData();
            bindingNavigator1.BindingSource = bindingSource1;

            bindingNavigator1.CausesValidation = true;

        }
        private void LoadData()
        {
            #region setup the binding

            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT * FROM Batteries ORDER BY BatterySerialNumber ASC";

            Bats.Clear();
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
                    myDataAdapter.Fill(Bats);
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
            bindingSource1.DataSource = null;
            bindingSource1.DataSource = Bats;

            bindingSource1.DataMember = "Table";

            #endregion

            #region setup the combo boxes
            ComboBox SerNumCB = toolStripCBSerNum.ComboBox;
            SerNumCB.DisplayMember = "BatterySerialNumber";
            SerNumCB.DataSource = bindingSource1;
            

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

            List<string> Customers =  Custs.Tables[0].AsEnumerable().Select(x => x[0].ToString()).Distinct().ToList();
            Customers.Sort();
            Customers.Insert(0, "");
            ComboBox CustCB = toolStripCBCustomers.ComboBox;
            //SerNumCB.DisplayMember = "BatterySerialNumber";
            CustCB.DataSource = Customers;

            //  Finally, setup the drop down to contain all customers availible in the customer table

            // Open database containing all the customer names data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT BatteryModel FROM BatteriesCustom ORDER BY BatteryModel ASC";

            DataSet BatsList = new DataSet();
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
                    myDataAdapter.Fill(BatsList);
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

            List<string> Mods = BatsList.Tables[0].AsEnumerable().Select(x => x[0].ToString()).Distinct().ToList();
            Mods.Sort();
            Mods.Insert(0, "");
            ComboBox ModCB = toolStripCBBatMod.ComboBox;
            //SerNumCB.DisplayMember = "BatterySerialNumber";
            ModCB.DataSource = Mods;
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

        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripCBCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateBats();
        }

        private void UpdateBats()
        {
            #region setup the binding

            // The xml to bind to.
            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....
            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";

            if (toolStripCBCustomers.Text == "" && toolStripCBBatMod.Text == "")
            {
                strAccessSelect = @"SELECT * FROM Batteries ORDER BY BatterySerialNumber ASC";
            }

            else if (toolStripCBBatMod.Text == "")
            {
                strAccessSelect = @"SELECT * FROM Batteries WHERE CustomerName='" + toolStripCBCustomers.Text.Replace("'", "''") + "' ORDER BY BatterySerialNumber ASC";
            }
            else if (toolStripCBCustomers.Text == "")
            {
                strAccessSelect = @"SELECT * FROM Batteries WHERE BatteryModel='" + toolStripCBBatMod.Text.Replace("'", "''") + "' ORDER BY BatterySerialNumber ASC";
            }
            else
            {
                strAccessSelect = @"SELECT * FROM Batteries WHERE CustomerName='" + toolStripCBCustomers.Text.Replace("'", "''") + "' AND " + "BatteryModel='" + toolStripCBBatMod.Text.Replace("'", "''") + "' ORDER BY BatterySerialNumber ASC";
            }

            Bats.Clear();
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
                    myDataAdapter.Fill(Bats);
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


            #endregion

            #region setup the combo boxes
            ComboBox SerNumCB = toolStripCBSerNum.ComboBox;
            SerNumCB.DisplayMember = "BatterySerialNumber";
            SerNumCB.DataSource = bindingSource1;

            #endregion
        }

        private void toolStripCBBatMod_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateBats();
        }

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {

        }

        private void bindingNavigator1_Layout(object sender, LayoutEventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void bindingNavigator1_TextChanged(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorMovePreviousItem_Click(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorMoveFirstItem_Click(object sender, EventArgs e)
        {

        }

        private void toolStripCBSerNum_TextChanged(object sender, EventArgs e)
        {

        }


        private void bindingNavigator1_ItemClicked_1(object sender, ToolStripItemClickedEventArgs e)
        {
        }

        private void frmVECustomerBats_FormClosing_1(object sender, FormClosingEventArgs e)
        {

        }

        private void frmVECustomerBats_Shown(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorMoveNextItem_Click(object sender, EventArgs e)
        {

        }

        private void bindingNavigator1_Validating_1(object sender, CancelEventArgs e)
        {

        }

        private void toolStripCBSerNum_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Load the new report based on the selected serial number
            if (toolStripCBSerNum.Text == "") { return; }

            // we need a data set..
            DataSet BatteryInfo = new DataSet();
            // Open database containing all the battery data....
            string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            string strAccessSelect = @"SELECT * FROM WorkOrders WHERE BatterySerialNumber='" + toolStripCBSerNum.Text + @"'";

            //Here is where I load the form wide dataset which will both let me fill in the rest of the combo boxes and the graphs!
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
                    myDataAdapter.Fill(BatteryInfo, "BatInfo");
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


            // Now that we have the data in WaterLevels lets pass it over to the matching report
            /*************************Load reportSet into reportSet  ************************/

            // bind datatable to report viewer
            this.reportViewer1.Reset();
            this.reportViewer1.ProcessingMode = ProcessingMode.Local;

            this.reportViewer1.LocalReport.ReportEmbeddedResource = "NewBTASProto.Reports.Battery_Report.rdlc";
            this.reportViewer1.LocalReport.DataSources.Clear();
            this.reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("WorkOrders", BatteryInfo.Tables[0]));

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



        }

        private void toolStripCBSerNum_Validating(object sender, CancelEventArgs e)
        {

        }

        private void toolStripCBSerNum_DropDown(object sender, EventArgs e)
        {
            
        }

        private void toolStripCBCustomers_DropDown(object sender, EventArgs e)
        {
            
        }

        private void toolStripCBBatMod_DropDown(object sender, EventArgs e)
        {
            
        }

        private void toolStripCBCustomers_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void toolStripCBBatMod_TextChanged(object sender, EventArgs e)
        {

        }

        private void toolStripCBCustomers_Enter(object sender, EventArgs e)
        {

        }

        private void toolStripCBCustomers_Leave(object sender, EventArgs e)
        {

        }

        private void toolStripCBBatMod_Enter(object sender, EventArgs e)
        {

        }

        private void toolStripCBBatMod_Leave(object sender, EventArgs e)
        {

        }

        private void toolStripCBSerNum_Enter(object sender, EventArgs e)
        {

        }

        private void Battery_Reports_Load(object sender, EventArgs e)
        {

            this.reportViewer1.RefreshReport();
        }
    }
}
