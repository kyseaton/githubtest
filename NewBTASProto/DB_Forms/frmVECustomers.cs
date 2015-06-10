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

        public frmVECustomers()
        {
            InitializeComponent();
            LoadData();
            bindingNavigator1.BindingSource = bindingSource1;
        }
        private void LoadData()
        {
            #region setup the binding

            // The xml to bind to.
            string strAccessConn;
            string strAccessSelect;
            // Open database containing all the battery data....

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
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
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }
            //  now try to access it
            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(Customers);

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

                    if (current["CustomerID"].ToString() != "")
                    {
                        //record already exist as we need to do an update

                        string cmdStr = "DELETE FROM CUSTOMERS WHERE CustomerID=" + current["CustomerID"].ToString();
                        OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                        cmd.ExecuteNonQuery();

                        // Also update the binding source
                        Customers.Tables[0].Rows[bindingNavigator1.BindingSource.Position].Delete();

                    }
                    else
                    {
                        MessageBox.Show("That record was not in the DB. You must save it in order to delete it.");
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Deletion Error" + ex.ToString());
            }
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
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
                    cmd.ExecuteNonQuery();

                }
                else
                {
                    // we need to insert a new record...
                    // find the max value in the CustomerID column so we know what to assign to the new record
                    int max = Customers.Tables[0].AsEnumerable().Max(r => r.Field<int>("CustomerID"));
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
                    cmd.ExecuteNonQuery();

                    // update the dataTable with the new customer ID also..
                    current[0] = max + 1;


                }
            }// end try
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {

        }
    }
}
