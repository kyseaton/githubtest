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
    public partial class frmVETechs : Form
    {

        DataSet Operators;

        public frmVETechs()
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

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
            strAccessSelect = @"SELECT * FROM Operators ORDER BY OperatorName ASC";

            Operators = new DataSet();
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
                myDataAdapter.Fill(Operators);

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
            bindingSource1.DataSource = Operators;
            bindingSource1.DataMember = "Table";

            textBox1.DataBindings.Add("Text", bindingSource1, "OperatorName");

            #endregion

            #region setup the combo box
            ComboBox CustomerCB = toolStripCBTechs.ComboBox;
            CustomerCB.DisplayMember = "OperatorName";
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
                if (MessageBox.Show("Are you sure you want to remove this Technician?", "Delete Record", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    // set up the db Connection
                    string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                    OleDbConnection conn = new OleDbConnection(connectionString);
                    conn.Open();

                    //get the current row
                    DataRowView current = (DataRowView)bindingSource1.Current;

                    // first test to see if the record already is actually in the database

                    if (current["ID"].ToString() != "")
                    {
                        //record already exist as we need to do an update

                        string cmdStr = "DELETE FROM Operators WHERE ID=" + current["ID"].ToString();
                        OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                        cmd.ExecuteNonQuery();

                        // Also update the binding source
                        Operators.Tables[0].Rows[bindingNavigator1.BindingSource.Position].Delete();

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
            if (bindingNavigator1.BindingSource.Position == -1)
            {
                string temp1 = textBox1.Text;

                bindingNavigator1.BindingSource.AddNew();
                bindingNavigator1.BindingSource.Position = 0;

                textBox1.Text = temp1;
            }

            try
            {

                // set up the db Connection
                string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                OleDbConnection conn = new OleDbConnection(connectionString);
                conn.Open();

                //MAKE SURE YOU SELECT THE CURRENT ROW FOR DOUBLE SAVES!!!!!!!!!!!!!!!!!

                //get the current row
                DataRowView current = (DataRowView)bindingSource1.Current;

                // first test to see if the record already is in the database

                //string cmdStr = "Select count(*) from CUSTOMERS where CustomerID=" + current["CustomerID"].ToString(); //get the existence of the record as count
                //OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                //int count = (int)cmd.ExecuteScalar();

                if (current["ID"].ToString() != "")
                {
                    //record already exist as we need to do an update

                    string cmdStr = "UPDATE Operators SET OperatorName='" + textBox1.Text.Replace("'", "''") +
                        "' WHERE ID=" + current["ID"].ToString();
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    cmd.ExecuteNonQuery();

                    //now force an update on the binding by moving one ahead and then back...
                    toolStripCBTechs.ComboBox.Text = textBox1.Text.Replace("'", "''");
                    MessageBox.Show(textBox1.Text.Replace("'", "''") + "'s entry has been updated.");

                }
                else
                {
                    // we need to insert a new record...
                    // find the max value in the CustomerID column so we know what to assign to the new record
                    int max;
                    try
                    {
                        max = Operators.Tables[0].AsEnumerable().Max(r => r.Field<int>("ID"));
                    }
                    catch
                    {
                        max = 0;
                    }
                    string cmdStr = "INSERT INTO Operators (ID, OperatorName) " +
                        "VALUES (" + (max + 1).ToString() + ",'" +
                        textBox1.Text + "')";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show(textBox1.Text + " has been been added to the operator list.");

                    // update the dataTable with the new customer ID also..
                    current[0] = max + 1; 
                }

                bindingNavigatorAddNewItem.Enabled = true;
            }// end try
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void frmVETechs_FormClosed(object sender, FormClosedEventArgs e)
        {
            ((Main_Form)this.Owner).Initialize_Operators_CB();
        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            bindingNavigatorAddNewItem.Enabled = false;
        }

        private void bindingNavigatorMovePreviousItem_Click(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false)
                {
                    Operators.Tables[0].Rows[Operators.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void bindingNavigatorMoveFirstItem_Click(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false)
                {
                    Operators.Tables[0].Rows[Operators.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void toolStripCBTechs_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false)
                {
                    Operators.Tables[0].Rows[Operators.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }
        }
    }
}
