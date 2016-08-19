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
    public partial class frmVEComboTests : Form
    {

        DataSet ComboTests;

        public frmVEComboTests()
        {
            InitializeComponent();
            LoadData();
            bindingNavigator1.BindingSource = bindingSource1;
            bindingNavigator1.CausesValidation = true;
        }
        private void LoadData()
        {

            string strAccessConn;
            string strAccessSelect;

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT TESTNAME FROM TestType ORDER BY TESTNAME ASC";

            DataSet Tests = new DataSet();
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

            List<string> SerialNums = Tests.Tables[0].AsEnumerable().Select(x => x[0].ToString()).Distinct().ToList();
            SerialNums.Sort();
            SerialNums.Insert(SerialNums.Count,"Wait");
            SerialNums.Insert(SerialNums.Count,"Temp Settle");

            foreach (string x in SerialNums)
            {
                comboBox1.Items.Add(x);
                comboBox2.Items.Add(x);
                comboBox3.Items.Add(x);
                comboBox4.Items.Add(x);
                comboBox5.Items.Add(x);
                comboBox6.Items.Add(x);
                comboBox7.Items.Add(x);
                comboBox8.Items.Add(x);
                comboBox9.Items.Add(x);
                comboBox10.Items.Add(x);
                comboBox11.Items.Add(x);
                comboBox12.Items.Add(x);
                comboBox13.Items.Add(x);
                comboBox14.Items.Add(x);
                comboBox15.Items.Add(x);
            }

            #region now setup the binding


            // Open database containing all the battery data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT * FROM ComboTest ORDER BY ComboTestName ASC";

            ComboTests = new DataSet();
            myAccessConn = null;
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
                    myDataAdapter.Fill(ComboTests);
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
            bindingSource1.DataSource = ComboTests;

            bindingSource1.DataMember = "Table";

            textBox1.DataBindings.Add("Text", bindingSource1, "ComboTestName");
            numericUpDown1.DataBindings.Add("Text", bindingSource1, "Steps");
            comboBox1.DataBindings.Add("Text", bindingSource1, "Step1");
            comboBox2.DataBindings.Add("Text", bindingSource1, "Step2");
            comboBox3.DataBindings.Add("Text", bindingSource1, "Step3");
            comboBox4.DataBindings.Add("Text", bindingSource1, "Step4");
            comboBox5.DataBindings.Add("Text", bindingSource1, "Step5");
            comboBox6.DataBindings.Add("Text", bindingSource1, "Step6");
            comboBox7.DataBindings.Add("Text", bindingSource1, "Step7");
            comboBox8.DataBindings.Add("Text", bindingSource1, "Step8");
            comboBox9.DataBindings.Add("Text", bindingSource1, "Step9");
            comboBox10.DataBindings.Add("Text", bindingSource1, "Step10");
            comboBox11.DataBindings.Add("Text", bindingSource1, "Step11");
            comboBox12.DataBindings.Add("Text", bindingSource1, "Step12");
            comboBox13.DataBindings.Add("Text", bindingSource1, "Step13");
            comboBox14.DataBindings.Add("Text", bindingSource1, "Step14");
            comboBox15.DataBindings.Add("Text", bindingSource1, "Step15");
            numericUpDown2.DataBindings.Add("Text", bindingSource1, "WaitTime");
            numericUpDown3.DataBindings.Add("Text", bindingSource1, "TempMar");

            #endregion

            #region setup the combo box
            ComboBox CustomerCB = toolStripCBTests.ComboBox;
            CustomerCB.DisplayMember = "ComboTestName";
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
            if (numericUpDown2.Text == "")
            {
                numericUpDown2.Text = "1";
            }
            if (numericUpDown3.Text == "")
            {
                numericUpDown3.Text = "1";
            }

            if (bindingNavigator1.BindingSource.Position == -1)
            {
                string temp1 = textBox1.Text;
                string temp2 = numericUpDown1.Text;
                string temp3 = comboBox1.Text;
                string temp4 = comboBox2.Text;
                string temp5 = comboBox3.Text;
                string temp6 = comboBox4.Text;
                string temp7 = comboBox5.Text;
                string temp8 = comboBox6.Text;
                string temp9 = comboBox7.Text;
                string temp10 = comboBox8.Text;
                string temp11 = comboBox9.Text;
                string temp12 = comboBox10.Text;
                string temp13 = comboBox11.Text;
                string temp14 = comboBox12.Text;
                string temp15 = comboBox13.Text;
                string temp16 = comboBox14.Text;
                string temp17 = comboBox15.Text;
                string temp18 = numericUpDown2.Text;
                string temp19 = numericUpDown3.Text;

                bindingNavigator1.BindingSource.AddNew();
                bindingNavigator1.BindingSource.Position = 0;

                textBox1.Text = temp1;
                numericUpDown1.Text = temp2;
                comboBox1.Text = temp3;
                comboBox2.Text = temp4;
                comboBox3.Text = temp5;
                comboBox4.Text = temp6;
                comboBox5.Text = temp7;
                comboBox6.Text = temp8;
                comboBox7.Text = temp9;
                comboBox8.Text = temp10;
                comboBox9.Text = temp11;
                comboBox10.Text = temp12;
                comboBox11.Text = temp13;
                comboBox12.Text = temp14;
                comboBox13.Text = temp15;
                comboBox14.Text = temp16;
                comboBox15.Text = temp17;
                numericUpDown2.Text = temp18;
                numericUpDown3.Text = temp19;

            }

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

                if (current["CTID"].ToString() != "")
                {
                    //record already exist as we need to do an update
                    string cmdStr = "UPDATE ComboTest SET ComboTestName = '" + textBox1.Text.Replace("'", "''") +
                        "', Steps=" + numericUpDown1.Text.Replace("'", "''") +
                        ", Step1='" + comboBox1.Text.Replace("'", "''") +
                        "', Step2='" + comboBox2.Text.Replace("'", "''") +
                        "', Step3='" + comboBox3.Text.Replace("'", "''") +
                        "', Step4='" + comboBox4.Text.Replace("'", "''") +
                        "', Step5='" + comboBox5.Text.Replace("'", "''") +
                        "', Step6='" + comboBox6.Text.Replace("'", "''") +
                        "', Step7='" + comboBox7.Text.Replace("'", "''") +
                        "', Step8='" + comboBox8.Text.Replace("'", "''") +
                        "', Step9='" + comboBox9.Text.Replace("'", "''") +
                        "', Step10='" + comboBox10.Text.Replace("'", "''") +
                        "', Step11='" + comboBox11.Text.Replace("'", "''") +
                        "', Step12='" + comboBox12.Text.Replace("'", "''") +
                        "', Step13='" + comboBox13.Text.Replace("'", "''") +
                        "', Step14='" + comboBox14.Text.Replace("'", "''") +
                        "', Step15='" + comboBox15.Text.Replace("'", "''") +
                        "', WaitTime='" + numericUpDown2.Text.Replace("'", "''") +
                        "', TempMar='" + numericUpDown3.Text.Replace("'", "''") +
                        "' WHERE CTID=" + current["CTID"].ToString();

                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    MessageBox.Show(this, current["ComboTestName"].ToString() + " has been updated.", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {

                    //somehow we are renamed the test and that's not good
                    string cmdStr = "INSERT INTO ComboTest (ComboTestName, Steps, Step1, Step2, Step3, Step4, Step5, Step6, Step7, Step8, Step9, Step10, Step11, Step12, Step13, Step14, Step15, WaitTime, TempMar) VALUES('" +
                        textBox1.Text.Replace("'","''") + "'," +
                        numericUpDown1.Text.Replace("'", "''") + ",'" +
                        comboBox1.Text.Replace("'", "''") + "','" +
                        comboBox2.Text.Replace("'","''") + "','" +
                        comboBox3.Text.Replace("'","''") + "','" +
                        comboBox4.Text.Replace("'","''") + "','" +
                        comboBox5.Text.Replace("'","''") + "','" +
                        comboBox6.Text.Replace("'","''") + "','" +
                        comboBox7.Text.Replace("'","''") + "','" +
                        comboBox8.Text.Replace("'","''") + "','" +
                        comboBox9.Text.Replace("'","''") + "','" +
                        comboBox10.Text.Replace("'","''") + "','" +
                        comboBox11.Text.Replace("'","''") + "','" +
                        comboBox12.Text.Replace("'","''") + "','" +
                        comboBox13.Text.Replace("'","''") + "','" +
                        comboBox14.Text.Replace("'","''") + "','" +
                        comboBox15.Text.Replace("'", "''") + "','" +
                        numericUpDown2.Text.Replace("'", "''") + "','" +
                        numericUpDown3.Text.Replace("'", "''") + "')";

                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    MessageBox.Show(this, current["ComboTestName"].ToString() + " has been created.", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

/*                if (bindingNavigator1.BindingSource.Position > 0)
                {
                    bindingNavigator1.BindingSource.Position -= 1;
                    bindingNavigator1.BindingSource.Position += 1;
                }
                else
                {
                    bindingNavigator1.BindingSource.Position += 1;
                    bindingNavigator1.BindingSource.Position -= 1;
                }
                */

                bindingNavigatorAddNewItem.Enabled = true;

                
            }// end try
            catch(Exception ex)
            {
                MessageBox.Show(this, ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

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
            if (bindingNavigatorAddNewItem.Enabled == false && bindingNavigator1.BindingSource.Position < ComboTests.Tables[0].Rows.Count)
            {
                ComboTests.Tables[0].Rows[ComboTests.Tables[0].Rows.Count - 1].Delete();
                bindingNavigatorAddNewItem.Enabled = true;
            }
        }

        private void bindingNavigatorMovePreviousItem_Click(object sender, EventArgs e)
        {
            //remove the new record if there is one..
            if (bindingNavigatorAddNewItem.Enabled == false && bindingNavigator1.BindingSource.Position < ComboTests.Tables[0].Rows.Count)
            {
                ComboTests.Tables[0].Rows[ComboTests.Tables[0].Rows.Count - 1].Delete();
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
                        string cmdStr = "DELETE FROM ComboTest WHERE CTID=" + current["CTID"].ToString();
                        OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                        lock (Main_Form.dataBaseLock)
                        {
                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }

                        // Also update the binding source
                        ComboTests.Tables[0].Rows[bindingNavigator1.BindingSource.Position].Delete();

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
            ((Main_Form)this.Owner).updateComboTestDropDown();
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

            if (numericUpDown1.Value >= 1) { comboBox1.Enabled = true; }
            else { comboBox1.Enabled = false; }
            if (numericUpDown1.Value >= 2) { comboBox2.Enabled = true; }
            else { comboBox2.Enabled = false; }
            if (numericUpDown1.Value >= 3) { comboBox3.Enabled = true; }
            else { comboBox3.Enabled = false; }
            if (numericUpDown1.Value >= 4) { comboBox4.Enabled = true; }
            else { comboBox4.Enabled = false; }
            if (numericUpDown1.Value >= 5) { comboBox5.Enabled = true; }
            else { comboBox5.Enabled = false; }
            if (numericUpDown1.Value >= 6) { comboBox6.Enabled = true; }
            else { comboBox6.Enabled = false; }
            if (numericUpDown1.Value >= 7) { comboBox7.Enabled = true; }
            else { comboBox7.Enabled = false; }
            if (numericUpDown1.Value >= 8) { comboBox8.Enabled = true; }
            else { comboBox8.Enabled = false; }
            if (numericUpDown1.Value >= 9) { comboBox9.Enabled = true; }
            else { comboBox9.Enabled = false; }
            if (numericUpDown1.Value >= 10) { comboBox10.Enabled = true; }
            else { comboBox10.Enabled = false; }
            if (numericUpDown1.Value >= 11) { comboBox11.Enabled = true; }
            else { comboBox11.Enabled = false; }
            if (numericUpDown1.Value >= 12) { comboBox12.Enabled = true; }
            else { comboBox12.Enabled = false; }
            if (numericUpDown1.Value >= 13) { comboBox13.Enabled = true; }
            else { comboBox13.Enabled = false; }
            if (numericUpDown1.Value >= 14) { comboBox14.Enabled = true; }
            else { comboBox14.Enabled = false; }
            if (numericUpDown1.Value >= 15) { comboBox15.Enabled = true; }
            else { comboBox15.Enabled = false; }

        }

    }
}
