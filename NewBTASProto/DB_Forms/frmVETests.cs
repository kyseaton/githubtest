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
    public partial class frmVETests : Form
    {

        DataSet Tests;

        public frmVETests()
        {
            InitializeComponent();
            LoadData();
            bindingNavigator1.BindingSource = bindingSource1;
            bindingNavigator1.CausesValidation = true;
        }
        private void LoadData()
        {
            #region setup the binding

            // The xml to bind to.
            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME<>'Top Charge-4' AND TESTNAME<>'As Received' AND TESTNAME<>'Full Charge-4' AND TESTNAME<>'Full Charge-4.5' AND TESTNAME<>'Full Charge-6' AND TESTNAME<>'Capacity-1' AND TESTNAME<>'Top Charge-2' AND TESTNAME<>'Discharge' AND TESTNAME<>'Slow Charge-14' AND TESTNAME<>'Top Charge-1' AND TESTNAME<>'Slow Charge-16' AND TESTNAME<>'Constant Voltage' AND TESTNAME<>'Shorting-16' ORDER BY TESTNAME ASC";

            Tests = new DataSet();
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
                    myDataAdapter.Fill(Tests);
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




            // Set the DataSource to the DataSet, and the DataMember
            bindingSource1.DataSource = null;
            bindingSource1.DataSource = Tests;

            bindingSource1.DataMember = "Table";

            textBox1.DataBindings.Add("Text", bindingSource1, "TESTNAME");
            numericUpDown1.DataBindings.Add("Text", bindingSource1, "Readings");
            numericUpDown2.DataBindings.Add("Text", bindingSource1, "Interval");
            comboBox2.DataBindings.Add("Text", bindingSource1, "TMode");
            numericUpDown6.DataBindings.Add("Text", bindingSource1, "TTime1Hr");
            numericUpDown5.DataBindings.Add("Text", bindingSource1, "TTime1Min");
            numericUpDown3.DataBindings.Add("Text", bindingSource1, "TCurr1");
            numericUpDown4.DataBindings.Add("Text", bindingSource1, "TVolts1");
            numericUpDown10.DataBindings.Add("Text", bindingSource1, "TTime2Hr");
            numericUpDown9.DataBindings.Add("Text", bindingSource1, "TTime2Min");
            numericUpDown8.DataBindings.Add("Text", bindingSource1, "TCurr2");
            numericUpDown7.DataBindings.Add("Text", bindingSource1, "TVolts2");
            numericUpDown13.DataBindings.Add("Text", bindingSource1, "TOhms");
            numericUpDown15.DataBindings.Add("Text", bindingSource1, "TTimeDHr");
            numericUpDown14.DataBindings.Add("Text", bindingSource1, "TTimeDMin");
            numericUpDown12.DataBindings.Add("Text", bindingSource1, "TCurrD");
            numericUpDown11.DataBindings.Add("Text", bindingSource1, "TVoltsD");


            #endregion

            #region setup the combo box
            ComboBox CustomerCB = toolStripCBTests.ComboBox;
            CustomerCB.DisplayMember = "TESTNAME";
            CustomerCB.DataSource = bindingSource1;


            #endregion

        }


        private void textBox5_Leave(object sender, EventArgs e)
        {

        }

        private void toolStripCBCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void saveToolStripButton_Click_1(object sender, EventArgs e)
        {

            try
            {

                // set up the db Connection
                string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                OleDbConnection conn = new OleDbConnection(connectionString);

                //MAKE SURE YOU SELECT THE CURRENT ROW FOR DOUBLE SAVES!!!!!!!!!!!!!!!!!

                //get the current row
                DataRowView current = (DataRowView)bindingSource1.Current;

                // first test to see if the record already is in the database

                //string cmdStr = "Select count(*) from CUSTOMERS where CustomerID=" + current["CustomerID"].ToString(); //get the existence of the record as count
                //OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                //int count = (int)cmd.ExecuteScalar();

                if (current["TESTID"].ToString() != "")
                {
                    //record already exist as we need to do an update
                    string cmdStr = "UPDATE TestType SET TESTNAME ='" + textBox1.Text + 
                        "', Readings='" + numericUpDown1.Text +
                        "', [Interval]='" + numericUpDown2.Text.Trim() +
                        "', TMode='" + comboBox2.Text +
                        "', TTime1Hr='" + numericUpDown6.Text.Trim() +
                        "', TTime1Min='" + numericUpDown5.Text.Trim() +
                        "', TCurr1='" + numericUpDown3.Text.Trim() +
                        "', TVolts1='" + numericUpDown4.Text.Trim() +
                        "', TTime2Hr='" + numericUpDown10.Text.Trim() +
                        "', TTime2Min='" + numericUpDown9.Text.Trim() +
                        "', TCurr2='" + numericUpDown8.Text.Trim() +
                        "', TVolts2='" + numericUpDown7.Text.Trim() +
                        "', TOhms='" + numericUpDown13.Text.Trim() +
                        "', TTimeDHr='" + numericUpDown15.Text.Trim() +
                        "', TTimeDMin='" + numericUpDown14.Text.Trim() +
                        "', TCurrD='" + numericUpDown12.Text.Trim() +
                        "', TVoltsD='" + numericUpDown11.Text.Trim() +
                        "' WHERE TESTID=" + current["TESTID"].ToString();

                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    MessageBox.Show(this, current["TESTNAME"].ToString() + " has been updated.", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    if (textBox1.Text == "As Received" ||
                        textBox1.Text == "Top Charge-4" ||
                        textBox1.Text == "Full Charge-4" ||
                        textBox1.Text == "Full Charge-6" ||
                        textBox1.Text == "Capacity-1" ||
                        textBox1.Text == "Top Charge-2" ||
                        textBox1.Text == "Discharge" ||
                        textBox1.Text == "Slow Charge-14" ||
                        textBox1.Text == "Top Charge-1" ||
                        textBox1.Text == "Slow Charge-16" ||
                        textBox1.Text == "Constant Voltage" ||
                        textBox1.Text == "Custom Chg" ||
                        textBox1.Text == "Custom Cap" ||
                        textBox1.Text == "Full Charge-4.5" ||
                        textBox1.Text == "Shorting-16")
                    {
                        MessageBox.Show(this, "That test name is protected.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }

                    string cmdStr = "INSERT INTO TestType (TESTNAME, Readings, [Interval], TMode, TTime1Hr, TTime1Min, TCurr1, TVolts1, TTime2Hr, TTime2Min, TCurr2, TVolts2, TOhms, TTimeDHr, TTimeDMin, TCurrD, TVoltsD) VALUES('" 
                        + textBox1.Text.Replace("'","''") + "','" 
                        + numericUpDown1.Text + "','"
                        + numericUpDown2.Text.Trim() + "','"
                        + comboBox2.Text + "','"
                        + numericUpDown6.Text + "','"
                        + numericUpDown5.Text + "','"
                        + numericUpDown3.Text + "','"
                        + numericUpDown4.Text + "','"
                        + numericUpDown10.Text + "','"
                        + numericUpDown9.Text + "','"
                        + numericUpDown8.Text + "','"
                        + numericUpDown7.Text + "','"
                        + numericUpDown13.Text + "','"
                        + numericUpDown15.Text + "','"
                        + numericUpDown14.Text + "','"
                        + numericUpDown12.Text + "','"
                        + numericUpDown11.Text
                        + "')";

                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    MessageBox.Show(this, current["TESTNAME"].ToString() + " has been created.", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                if (bindingNavigator1.BindingSource.Position > 0)
                {
                    bindingNavigator1.BindingSource.Position -= 1;
                    bindingNavigator1.BindingSource.Position += 1;
                }
                else
                {
                    bindingNavigator1.BindingSource.Position += 1;
                    bindingNavigator1.BindingSource.Position -= 1;
                }

                bindingNavigatorAddNewItem.Enabled = true;

                
            }// end try
            catch(Exception ex)
            {
                MessageBox.Show(this, ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "Custom Cap" || textBox1.Text == "Custom Chg")
            {
                textBox1.Enabled = false;
                bindingNavigatorDeleteItem.Enabled = false;
                label15.Visible = true;
                comboBox2.Enabled = false;
                numericUpDown6.Enabled = false;
                numericUpDown5.Enabled = false;
                numericUpDown3.Enabled = false;
                numericUpDown4.Enabled = false;
                numericUpDown10.Enabled = false;
                numericUpDown9.Enabled = false;
                numericUpDown8.Enabled = false;
                numericUpDown7.Enabled = false;
                numericUpDown15.Enabled = false;
                numericUpDown14.Enabled = false;
                numericUpDown12.Enabled = false;
                numericUpDown11.Enabled = false;
                numericUpDown13.Enabled = false;

            }
            else
            {
                textBox1.Enabled = true;
                bindingNavigatorDeleteItem.Enabled = true;
                label15.Visible = false;
                comboBox2.Enabled = true;
                numericUpDown6.Enabled = true;
                numericUpDown5.Enabled = true;
                numericUpDown3.Enabled = true;
                numericUpDown4.Enabled = true;
                numericUpDown10.Enabled = true;
                numericUpDown9.Enabled = true;
                numericUpDown8.Enabled = true;
                numericUpDown7.Enabled = true;
                numericUpDown15.Enabled = true;
                numericUpDown14.Enabled = true;
                numericUpDown12.Enabled = true;
                numericUpDown11.Enabled = true;
                numericUpDown13.Enabled = true;
            }
        }

        private void toolStripCBTests_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (toolStripCBTests.Text == "")
            {
                bindingNavigatorAddNewItem.Enabled = false;
            }

        }

        private void bindingNavigatorMoveFirstItem_Click(object sender, EventArgs e)
        {
            //remove the new record if there is one..
            if (bindingNavigatorAddNewItem.Enabled == false && bindingNavigator1.BindingSource.Position < Tests.Tables[0].Rows.Count)
            {
                Tests.Tables[0].Rows[Tests.Tables[0].Rows.Count - 1].Delete();
                bindingNavigatorAddNewItem.Enabled = true;
            }
        }

        private void bindingNavigatorMovePreviousItem_Click(object sender, EventArgs e)
        {
            //remove the new record if there is one..
            if (bindingNavigatorAddNewItem.Enabled == false && bindingNavigator1.BindingSource.Position < Tests.Tables[0].Rows.Count)
            {
                Tests.Tables[0].Rows[Tests.Tables[0].Rows.Count - 1].Delete();
                bindingNavigatorAddNewItem.Enabled = true;
            }
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show(this, "Are you sure you want to remove this custom test?", "Delete Record", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    // set up the db Connection
                    string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                    OleDbConnection conn = new OleDbConnection(connectionString);

                    //get the current row
                    DataRowView current = (DataRowView)bindingSource1.Current;

                    // first test to see if the record already is actually in the database

                    if (toolStripCBTests.Text != "")
                    {              
                        // first delete the tests and scandata!
                        string cmdStr = "DELETE FROM TestType WHERE TESTNAME='" + current["TESTNAME"].ToString() + "'";
                        OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                        lock (Main_Form.dataBaseLock)
                        {
                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }

                        // Also update the binding source
                        Tests.Tables[0].Rows[bindingNavigator1.BindingSource.Position].Delete();

                    }
                    else
                    {
                        MessageBox.Show(this, "That record was not in the DB. You must save it in order to delete it.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Deletion Error" + ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void frmVETests_FormClosing(object sender, FormClosingEventArgs e)
        {
            ((Main_Form)this.Owner).updateCustomTestDropDown();
        }

    }
}
