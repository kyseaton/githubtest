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

namespace NewBTASProto
{
    public partial class frmVECustomBats : Form
    {

        DataSet CustomBats;

        public frmVECustomBats()
        {
            InitializeComponent();
            LoadData();
            bindingNavigator1.BindingSource = bindingSource1;
            bindingNavigator1.Select();
        }
        private void LoadData()
        {
            #region setup the binding

            // The xml to bind to.
            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
            strAccessSelect = @"SELECT * FROM BatteriesCustom ORDER BY BatteryModel ASC";

            CustomBats = new DataSet();
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
                myDataAdapter.Fill(CustomBats);

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




            // Set the DataSource to the DataSet, and the DataMember
            bindingSource1.DataSource = null;
            bindingSource1.DataSource = CustomBats;

            bindingSource1.DataMember = "Table";

            textBox1.DataBindings.Add("Text", bindingSource1, "BMFR");
            textBox2.DataBindings.Add("Text", bindingSource1, "BatteryModel");
            textBox3.DataBindings.Add("Text", bindingSource1, "BPN");
            textBox4.DataBindings.Add("Text", bindingSource1, "BTECH");
            textBox5.DataBindings.Add("Text", bindingSource1, "BAPP");
            textBox6.DataBindings.Add("Text", bindingSource1, "VOLT");
            textBox7.DataBindings.Add("Text", bindingSource1, "NCELLS");
            textBox8.DataBindings.Add("Text", bindingSource1, "CAP");
            textBox9.DataBindings.Add("Text", bindingSource1, "CELL");
            textBox10.DataBindings.Add("Text", bindingSource1, "CPN");
            textBox11.DataBindings.Add("Text", bindingSource1, "CTORQU");
            textBox12.DataBindings.Add("Text", bindingSource1, "MCC");
            textBox13.DataBindings.Add("Text", bindingSource1, "MCT");
            textBox14.DataBindings.Add("Text", bindingSource1, "MPV");
            textBox15.DataBindings.Add("Text", bindingSource1, "TCC");
            textBox16.DataBindings.Add("Text", bindingSource1, "TCT");
            textBox17.DataBindings.Add("Text", bindingSource1, "TPV");
            textBox18.DataBindings.Add("Text", bindingSource1, "SCC");
            textBox19.DataBindings.Add("Text", bindingSource1, "SCT");
            textBox20.DataBindings.Add("Text", bindingSource1, "SPV");
            textBox21.DataBindings.Add("Text", bindingSource1, "BCVMIN");
            textBox22.DataBindings.Add("Text", bindingSource1, "BCVMAX");
            textBox23.DataBindings.Add("Text", bindingSource1, "COT");
            textBox24.DataBindings.Add("Text", bindingSource1, "CTC");
            textBox25.DataBindings.Add("Text", bindingSource1, "CTT");
            textBox26.DataBindings.Add("Text", bindingSource1, "CTMV");
            textBox27.DataBindings.Add("Text", bindingSource1, "CCVMMIN");
            textBox28.DataBindings.Add("Text", bindingSource1, "CCVMAX");
            textBox29.DataBindings.Add("Text", bindingSource1, "CCAPV");
            textBox30.DataBindings.Add("Text", bindingSource1, "TS1");
            textBox31.DataBindings.Add("Text", bindingSource1, "TS2");
            textBox32.DataBindings.Add("Text", bindingSource1, "SLACV");
            textBox33.DataBindings.Add("Text", bindingSource1, "SLAPK");
            textBox34.DataBindings.Add("Text", bindingSource1, "SLACVC");
            textBox35.DataBindings.Add("Text", bindingSource1, "SLAPKC");
            textBox36.DataBindings.Add("Text", bindingSource1, "NOTES");

            #endregion

            #region setup the combo box
            ComboBox CustomerCB = toolStripCBCustomers.ComboBox;
            CustomerCB.DisplayMember = "BatteryModel";
            CustomerCB.DataSource = bindingSource1;


            #endregion

        }

        #region unused

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

        private void textBox5_Leave(object sender, EventArgs e)
        {

            // Maybe down the line...

            /*            ulong num;

                        if (textBox5.Text.Length == 11 && ulong.TryParse(textBox5.Text, out num))
                        {

                            string pn = textBox5.Text;

                            textBox5.Text = String.Format("({0}) {1}-{2}", pn.Substring(0, 3), pn.Substring(3, 3), pn.Substring(6));

                        }
                        if (textBox5.Text.Length == 11 && ulong.TryParse(textBox5.Text, out num))
                        {

                            string pn = textBox5.Text;

                            textBox5.Text = String.Format("({0}) {1}-{2}", pn.Substring(0, 3), pn.Substring(3, 3), pn.Substring(6));

                        }

                        else
                        {

                            MessageBox.Show("Invalid phone number, please change");

                            textBox5.Focus();


                        }
             *  * */
        }

        private void toolStripCBCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {
           //this didn't work...
           // bindingNavigator1.BindingSource.Position = toolStripCBCustomers.SelectedIndex;
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {

        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void label31_Click(object sender, EventArgs e)
        {

        }
        # endregion


        private void bindingNavigatorDeleteItem_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Are you sure you want to remove this customer?", "Delete Record", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    // set up the db Connection
                    string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                    OleDbConnection conn = new OleDbConnection(connectionString);
                    conn.Open();

                    //get the current row
                    DataRowView current = (DataRowView)bindingSource1.Current;

                    // first test to see if the record already is actually in the database

                    if (current["RecordID"].ToString() != "")
                    {
                        //record already exist as we need to do an update

                        string cmdStr = "DELETE FROM BatteriesCustom WHERE RecordID=" + current["RecordID"].ToString();
                        OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                        cmd.ExecuteNonQuery();

                        // Also update the binding source
                        CustomBats.Tables[0].Rows[bindingNavigator1.BindingSource.Position].Delete();

                    }
                    else
                    {
                        MessageBox.Show("That record was not in the DB. You must save it in order to delete it.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Deletion Error" + ex.ToString());
            }
        }

        private void saveToolStripButton_Click_1(object sender, EventArgs e)
        {
            try
            {

                // set up the db Connection
                string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                OleDbConnection conn = new OleDbConnection(connectionString);
                conn.Open();

                //MAKE SURE YOU SELECT THE CURRENT ROW FOR DOUBLE SAVES!!!!!!!!!!!!!!!!!

                //get the current row
                DataRowView current = (DataRowView)bindingSource1.Current;

                // first test to see if the record already is in the database

                //string cmdStr = "Select count(*) from CUSTOMERS where CustomerID=" + current["CustomerID"].ToString(); //get the existence of the record as count
                //OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                //int count = (int)cmd.ExecuteScalar();

                if (current["RecordID"].ToString() != "")
                {
                    //record already exist as we need to do an update

                    string cmdStr = "UPDATE BatteriesCustom SET BMFR='" + textBox1.Text.Replace("'", "''") +
                        "', BatteryModel='" + textBox2.Text.Replace("'", "''") +
                        "', BPN='" + textBox3.Text.Replace("'", "''") +
                        "', BTECH='" + textBox4.Text.Replace("'", "''") +
                        "', BAPP='" + textBox5.Text.Replace("'", "''") +
                        "', VOLT='" + textBox6.Text.Replace("'", "''") +
                        "', NCELLS='" + textBox7.Text.Replace("'", "''") +
                        "', CAP='" + textBox8.Text.Replace("'", "''") +
                        "', CELL='" + textBox9.Text.Replace("'", "''") +
                        "', CPN='" + textBox10.Text.Replace("'", "''") +
                        "', CTORQU='" + textBox11.Text.Replace("'", "''") +
                        "', MCC='" + textBox12.Text.Replace("'", "''") +
                        "', MCT='" + textBox13.Text.Replace("'", "''") +
                        "', MPV='" + textBox14.Text.Replace("'", "''") +
                        "', TCC='" + textBox15.Text.Replace("'", "''") +
                        "', TCT='" + textBox16.Text.Replace("'", "''") +
                        "', TPV='" + textBox17.Text.Replace("'", "''") +
                        "', SCC='" + textBox18.Text.Replace("'", "''") +
                        "', SCT='" + textBox19.Text.Replace("'", "''") +
                        "', SPV='" + textBox20.Text.Replace("'", "''") +
                        "', BCVMIN='" + textBox21.Text.Replace("'", "''") +
                        "', BCVMAX='" + textBox22.Text.Replace("'", "''") +
                        "', COT='" + textBox23.Text.Replace("'", "''") +
                        "', CTC='" + textBox24.Text.Replace("'", "''") +
                        "', CTT='" + textBox25.Text.Replace("'", "''") +
                        "', CTMV='" + textBox26.Text.Replace("'", "''") +
                        "', CCVMMIN='" + textBox27.Text.Replace("'", "''") +
                        "', CCVMAX='" + textBox28.Text.Replace("'", "''") +
                        "', CCAPV='" + textBox29.Text.Replace("'", "''") +
                        "', TS1='" + textBox30.Text.Replace("'", "''") +
                        "', TS2='" + textBox31.Text.Replace("'", "''") +
                        "', SLACV='" + textBox32.Text.Replace("'", "''") +
                        "', SLAPK='" + textBox33.Text.Replace("'", "''") +
                        "', SLACVC='" + textBox34.Text.Replace("'", "''") +
                        "', SLAPKC='" + textBox35.Text.Replace("'", "''") +
                        "', NOTES='" + textBox36.Text.Replace("'", "''") +
                        "' WHERE RecordID=" + current["RecordID"].ToString();
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    cmd.ExecuteNonQuery();

                }
                else
                {
                    // we need to insert a new record...
                    // find the max value in the CustomerID column so we know what to assign to the new record
                    int max = CustomBats.Tables[0].AsEnumerable().Max(r => r.Field<int>("RecordID"));
                    string cmdStr = "INSERT INTO BatteriesCustom (RecordID, BMFR, BatteryModel, BPN, BTECH, BAPP, VOLT, NCELLS, CAP, CELL, CPN, CTORQU, MCC, MCT, MPV, TCC, TCT, " +
                       "TPV, SCC, SCT, SPV, BCVMIN, BCVMAX, COT, CTC, CTT, CTMV, CCVMMIN, CCVMAX, CCAPV, TS1, TS2, SLACV, SLAPK, SLACVC, SLAPKC, NOTES) " +
                        "VALUES (" + (max + 1).ToString() + ",'" +
                        textBox1.Text.Replace("'", "''") + "','" +
                        textBox2.Text.Replace("'", "''") + "','" +
                        textBox3.Text.Replace("'", "''") + "','" +
                        textBox4.Text.Replace("'", "''") + "','" +
                        textBox5.Text.Replace("'", "''") + "','" +
                        textBox6.Text.Replace("'", "''") + "','" +
                        textBox7.Text.Replace("'", "''") + "','" +
                        textBox8.Text.Replace("'", "''") + "','" +
                        textBox9.Text.Replace("'", "''") + "','" +
                        textBox10.Text.Replace("'", "''") + "','" +
                        textBox11.Text.Replace("'", "''") + "','" +
                        textBox12.Text.Replace("'", "''") + "','" +
                        textBox13.Text.Replace("'", "''") + "','" +
                        textBox14.Text.Replace("'", "''") + "','" +
                        textBox15.Text.Replace("'", "''") + "','" +
                        textBox16.Text.Replace("'", "''") + "','" +
                        textBox17.Text.Replace("'", "''") + "','" +
                        textBox18.Text.Replace("'", "''") + "','" +
                        textBox19.Text.Replace("'", "''") + "','" +
                        textBox20.Text.Replace("'", "''") + "','" +
                        textBox21.Text.Replace("'", "''") + "','" +
                        textBox22.Text.Replace("'", "''") + "','" +
                        textBox23.Text.Replace("'", "''") + "','" +
                        textBox24.Text.Replace("'", "''") + "','" +
                        textBox25.Text.Replace("'", "''") + "','" +
                        textBox26.Text.Replace("'", "''") + "','" +
                        textBox27.Text.Replace("'", "''") + "','" +
                        textBox28.Text.Replace("'", "''") + "','" +
                        textBox29.Text.Replace("'", "''") + "','" +
                        textBox30.Text.Replace("'", "''") + "','" +
                        textBox31.Text.Replace("'", "''") + "','" +
                        textBox32.Text.Replace("'", "''") + "','" +
                        textBox33.Text.Replace("'", "''") + "','" +
                        textBox34.Text.Replace("'", "''") + "','" +
                        textBox35.Text.Replace("'", "''") + "','" +
                        textBox36.Text.Replace("'", "''") + "')";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    cmd.ExecuteNonQuery();

                    // update the dataTable with the new customer ID also..
                    current[0] = max + 1;


                }
            }// end try
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {

        }



    }
}
