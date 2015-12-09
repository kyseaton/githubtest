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
    public partial class frmVECustomers : Form
    {

        DataSet Customers;

        // we use this bool to allow us to allow the databinding indext to be changed...
        bool Inhibit = true;
        bool InhibitCB = true;

        //current data..
        string curTemp1;
        string curTemp2;
        string curTemp3;
        string curTemp4;
        string curTemp5;
        string curTemp6;
        string curTemp7;
        string curTemp8;
        

        public frmVECustomers()
        {
            InitializeComponent();
            LoadData();
            bindingNavigator1.BindingSource = bindingSource1;

            bindingNavigator1.CausesValidation = true;

            Inhibit = false;
            InhibitCB = false;
        }
        private void LoadData()
        {
            #region setup the binding

            // The xml to bind to.
            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT * FROM CUSTOMERS ORDER BY CustomerName ASC";

            Customers = new DataSet();
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
                    myDataAdapter.Fill(Customers);
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
            bindingSource1.DataSource = Customers;

            bindingSource1.DataMember = "Table";

            textBox1.DataBindings.Add("Text", bindingSource1, "CustomerName");
            textBox2.DataBindings.Add("Text", bindingSource1, "Address1");
            textBox3.DataBindings.Add("Text", bindingSource1, "Address2");
            textBox4.DataBindings.Add("Text", bindingSource1, "Address3");
            textBox5.DataBindings.Add("Text", bindingSource1, "Phone");
            textBox6.DataBindings.Add("Text", bindingSource1, "Fax");
            textBox7.DataBindings.Add("Text", bindingSource1, "Contact");
            textBox8.DataBindings.Add("Text", bindingSource1, "E-Mail");
            textBox9.DataBindings.Add("Text", bindingSource1, "Notes");

            #endregion

            #region setup the combo box
            ComboBox CustomerCB = toolStripCBCustomers.ComboBox;
            CustomerCB.DisplayMember = "CustomerName";
            CustomerCB.DataSource = bindingSource1;


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
            updateCurVals();
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

                            MessageBox.Show(new Form() { TopMost = true }, "Invalid phone number, please change");

                            textBox5.Focus();


                        }
             *  * */
        }

        string oldCustName = "";

        private void toolStripCBCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (InhibitCB) { return; }

            //Validate before moving
            if (ValidateIt())
            {
                InhibitCB = true;
                // move back
                toolStripCBCustomers.SelectedIndex = bindingNavigator1.BindingSource.Position;
                InhibitCB = false;

            }
            else
            {
                InhibitCB = false;
                updateCurVals();
                oldCustName = textBox1.Text;
            }
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show(new Form() { TopMost = true }, "Are you sure you want to remove this customer?", "Delete Record", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    // set up the db Connection
                    string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                    OleDbConnection conn = new OleDbConnection(connectionString);

                    //get the current row
                    DataRowView current = (DataRowView)bindingSource1.Current;

                    // first test to see if the record already is actually in the database

                    if (current["CustomerID"].ToString() != "")
                    {
                        //record already exist as we need to do an update

                        string cmdStr = "DELETE FROM CUSTOMERS WHERE CustomerID=" + current["CustomerID"].ToString();
                        OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                        lock (Main_Form.dataBaseLock)
                        {
                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }

                        // Also update the binding source
                        Customers.Tables[0].Rows[bindingNavigator1.BindingSource.Position].Delete();

                    }
                    else
                    {
                        MessageBox.Show(new Form() { TopMost = true }, "That record was not in the DB. You must save it in order to delete it.");
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(new Form() { TopMost = true }, "Deletion Error" + ex.ToString());
            }
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            if (bindingNavigator1.BindingSource.Position == -1)
            {
                string temp1 = textBox1.Text;
                string temp2 = textBox2.Text;
                string temp3 = textBox3.Text;
                string temp4 = textBox4.Text;
                string temp5 = textBox5.Text;
                string temp6 = textBox6.Text;
                string temp7 = textBox7.Text;
                string temp8 = textBox8.Text;

                bindingNavigator1.BindingSource.AddNew();
                bindingNavigator1.BindingSource.Position = 0;

                textBox1.Text = temp1;
                textBox2.Text = temp2;
                textBox3.Text = temp3;
                textBox4.Text = temp4;
                textBox5.Text = temp5;
                textBox6.Text = temp6;
                textBox7.Text = temp7;
                textBox8.Text = temp8;

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

                if (current["CustomerID"].ToString() != "")
                {
                    //record already exist as we need to do an update

                    string cmdStr = "UPDATE CUSTOMERS SET CustomerName='" + textBox1.Text.Replace("'","''") +
                        "', Address1='" + textBox2.Text.Replace("'", "''") +
                        "', Address2='" + textBox3.Text.Replace("'", "''") +
                        "', Address3='" + textBox4.Text.Replace("'", "''") +
                        "', Phone='" + textBox5.Text.Replace("'", "''") +
                        "', Fax='" + textBox6.Text.Replace("'", "''") +
                        "', Contact='" + textBox7.Text.Replace("'", "''") +
                        "', [E-Mail]='" + textBox8.Text.Replace("'", "''") +
                        "', Notes='" + textBox9.Text.Replace("'", "''") +
                        "' WHERE CustomerID=" + current["CustomerID"].ToString();
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                    // Also update the customer name in the other tables!
                    cmdStr = "UPDATE WorkOrders SET CustomerName='" + textBox1.Text.Replace("'", "''") + "' WHERE CustomerName='" + oldCustName + "'";
                    cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                    cmdStr = "UPDATE Batteries SET CustomerName='" + textBox1.Text.Replace("'", "''") + "' WHERE CustomerName='" + oldCustName + "'";
                    cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                    //now force an update on the binding by moving one ahead and then back...
                    toolStripCBCustomers.ComboBox.Text = textBox1.Text.Replace("'", "''");

                    MessageBox.Show(textBox1.Text.Replace("'", "''") + "'s entry has been updated.");

                }
                else
                {
                    // we need to insert a new record...
                    // find the max value in the CustomerID column so we know what to assign to the new record
                    int max;
                    try
                    {
                        max = Customers.Tables[0].AsEnumerable().Max(r => r.Field<int>("CustomerID"));
                    }
                    catch
                    {
                        max = 0;
                    }
                    string cmdStr = "INSERT INTO CUSTOMERS (CustomerID, CustomerName, Address1, Address2, Address3, Phone, Fax, Contact, [E-Mail], Notes) " +
                        "VALUES (" + (max + 1).ToString() + ",'" +
                        textBox1.Text.Replace("'", "''") + "','" +
                        textBox2.Text.Replace("'", "''") + "','" +
                        textBox3.Text.Replace("'", "''") + "','" +
                        textBox4.Text.Replace("'", "''") + "','" +
                        textBox5.Text.Replace("'", "''") + "','" +
                        textBox6.Text.Replace("'", "''") + "','" +
                        textBox7.Text.Replace("'", "''") + "','" +
                        textBox8.Text.Replace("'", "''") + "','" +
                        textBox9.Text.Replace("'", "''") + "')";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    MessageBox.Show(textBox1.Text + " has been added as a customer.");

                    // update the dataTable with the new customer ID also..
                    current[0] = max + 1;

                    bindingNavigatorAddNewItem.Enabled = true;
                    
                }
                updateCurVals();
            }// end try
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            if (toolStripCBCustomers.Text == "")
            {
                bindingNavigatorAddNewItem.Enabled = false;
            }
        }

        private void bindingNavigatorMoveFirstItem_Click(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && bindingNavigator1.BindingSource.Position < Customers.Tables[0].Rows.Count)
                {
                    Customers.Tables[0].Rows[Customers.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void bindingNavigatorMovePreviousItem_Click(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && bindingNavigator1.BindingSource.Position < Customers.Tables[0].Rows.Count)
                {
                    Customers.Tables[0].Rows[Customers.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void toolStripCBCustomers_TextChanged(object sender, EventArgs e)
        {

            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && bindingNavigator1.BindingSource.Position < Customers.Tables[0].Rows.Count)
                {
                    Customers.Tables[0].Rows[Customers.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void bindingNavigator1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            Inhibit = false;
            bindingNavigator1.Focus();
        }

        private void frmVECustomers_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Validate before moving
            if (ValidateIt())
            {
                Inhibit = true;
                // move back
                e.Cancel = true;
            }
            else
            {
                Inhibit = true;
                updateCurVals();
            }
        }

        private bool ValidateIt()
        {
            // do we need to validate?

            if ( curTemp1 != textBox1.Text ||
                curTemp2 != textBox2.Text ||
                curTemp3 != textBox3.Text ||
                curTemp4 != textBox4.Text ||
                curTemp5 != textBox5.Text ||
                curTemp6 != textBox6.Text ||
                curTemp7 != textBox7.Text ||
                curTemp8 != textBox8.Text)
            {
                // they don't match!
                // ask if the user is sure that they want to continue...
                DialogResult dialogResult = MessageBox.Show(this, "Looks like this record has been updated without being saved.  Are you sure you want to navigate away without saving?", "Click Yes to continue or No to stop the test.", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                {
                    return true;
                }
            }
            return false;
        }

        private void bindingNavigator1_Validating(object sender, CancelEventArgs e)
        {
            if (Inhibit) { return; }

            //Validate before moving
            if (ValidateIt())
            {
                Inhibit = true;
                // move back
                e.Cancel = true;

            }
            else
            {
                Inhibit = true;
                updateCurVals();
            }
        }

        private void frmVECustomers_Shown(object sender, EventArgs e)
        {
            updateCurVals();

            bindingNavigatorAddNewItem.PerformClick();
        }

        private void bindingNavigator1_LocationChanged(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && bindingNavigator1.BindingSource.Position < Customers.Tables[0].Rows.Count)
                {
                    Customers.Tables[0].Rows[Customers.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void updateCurVals()
        {
            // update the current vars....
            //current data..
            curTemp1 = textBox1.Text;
            curTemp2 = textBox2.Text;
            curTemp3 = textBox3.Text;
            curTemp4 = textBox4.Text;
            curTemp5 = textBox5.Text;
            curTemp6 = textBox6.Text;
            curTemp7 = textBox7.Text;
            curTemp8 = textBox8.Text;
        }
    }
}
