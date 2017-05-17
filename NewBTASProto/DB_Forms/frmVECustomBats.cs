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
using System.Runtime.InteropServices;




namespace NewBTASProto
{
    public partial class frmVECustomBats : Form
    {
        [DllImport("user32.dll")]
        static extern bool LockWindowUpdate(IntPtr hWndLock);

        DataSet CustomBats;


        //current data..
        string curTemp1;
        string curTemp2;
        string curTemp3;
        string curTemp4;

        string curTemp5;
        string curTemp6;
        string curTemp7;

        string curTemp10;
        string curTemp11;
        string curTemp12;

        string curTemp13;
        string curTemp14;
        string curTemp15;

        string curTemp16;

        string curTemp17;

        decimal curTemp18;
        decimal curTemp19;
        decimal curTemp20;
        decimal curTemp21;

        decimal curTemp22;
        decimal curTemp23;
        decimal curTemp24;
        decimal curTemp25;

        string curTemp26;

        decimal curTemp27;
        decimal curTemp28;
        decimal curTemp29;
        decimal curTemp30;

        decimal curTemp31;
        decimal curTemp32;
        decimal curTemp33;
        decimal curTemp34;

        string curTemp35;

        decimal curTemp36;
        decimal curTemp37;
        decimal curTemp38;
        decimal curTemp39;

        decimal curTemp40;
        decimal curTemp41;
        decimal curTemp42;
        decimal curTemp43;

        string curTemp44;

        decimal curTemp45;
        decimal curTemp46;
        decimal curTemp47;
        decimal curTemp48;

        decimal curTemp49;
        decimal curTemp50;
        decimal curTemp51;
        decimal curTemp52;

        string curTemp53;

        decimal curTemp54;
        decimal curTemp55;
        decimal curTemp56;
        decimal curTemp57;

        decimal curTemp58;
        decimal curTemp59;
        decimal curTemp60;
        decimal curTemp61;

        string curTemp62;

        decimal curTemp63;
        decimal curTemp64;
        decimal curTemp65;
        decimal curTemp66;

        decimal curTemp67;
        decimal curTemp68;
        decimal curTemp69;
        decimal curTemp70;

        string curTemp71;

        decimal curTemp72;
        decimal curTemp73;
        decimal curTemp74;
        decimal curTemp75;
        decimal curTemp76;

        string curTemp77;

        decimal curTemp78;
        decimal curTemp79;
        decimal curTemp80;
        decimal curTemp81;
        decimal curTemp82;

        string curTemp83;

        decimal curTemp84;
        decimal curTemp85;
        decimal curTemp86;
        decimal curTemp87;

        decimal curTemp88;
        decimal curTemp89;
        decimal curTemp90;
        decimal curTemp101;

        string curTemp91;

        decimal curTemp92;
        decimal curTemp93;
        decimal curTemp94;
        decimal curTemp95;

        decimal curTemp96;
        decimal curTemp97;
        decimal curTemp98;
        decimal curTemp99;

        string curTemp100;

        decimal curTemp102;
        decimal curTemp103;
        decimal curTemp104;
        decimal curTemp105;

        decimal curTemp106;
        decimal curTemp107;
        decimal curTemp108;
        decimal curTemp109;

        string curTemp110;

        decimal curTemp111;
        decimal curTemp112;
        decimal curTemp113;
        decimal curTemp114;
        decimal curTemp115;

        string curTemp116;

        decimal curTemp117;
        decimal curTemp118;
        decimal curTemp119;
        decimal curTemp120;

        string curTemp121;

        decimal curTemp122;
        decimal curTemp123;
        decimal curTemp124;
        decimal curTemp125;

        decimal curTemp126;
        decimal curTemp127;
        decimal curTemp128;
        decimal curTemp129;

        string curTemp130;

        decimal curTemp131;
        decimal curTemp132;
        decimal curTemp133;
        decimal curTemp134;

        decimal curTemp135;
        decimal curTemp136;
        decimal curTemp137;
        decimal curTemp138;

        public frmVECustomBats()
        {
            InitializeComponent();
            LoadData();
            bindingNavigator1.BindingSource = bindingSource1;
            bindingNavigator1.Select();
            SetupForms();
            
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

            strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
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
                    myDataAdapter.Fill(CustomBats);
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




            // Set the DataSource to the DataSet, and the DataMember
            bindingSource1.DataSource = null;
            bindingSource1.DataSource = CustomBats;

            bindingSource1.DataMember = "Table";

            textBox1.DataBindings.Add("Text", bindingSource1, "BMFR");
            textBox2.DataBindings.Add("Text", bindingSource1, "BatteryModel");
            textBox3.DataBindings.Add("Text", bindingSource1, "BPN");
            comboBox13.DataBindings.Add("Text", bindingSource1, "BTECH");
            numericUpDown89.DataBindings.Add("Text", bindingSource1, "VOLT");
            numericUpDown88.DataBindings.Add("Text", bindingSource1, "NCELLS");
            numericUpDown90.DataBindings.Add("Text", bindingSource1, "CAP");
            textBox21.DataBindings.Add("Text", bindingSource1, "BCVMIN");
            textBox22.DataBindings.Add("Text", bindingSource1, "BCVMAX");
            textBox23.DataBindings.Add("Text", bindingSource1, "COT");
            textBox27.DataBindings.Add("Text", bindingSource1, "CCVMMIN");
            textBox28.DataBindings.Add("Text", bindingSource1, "CCVMAX");
            textBox29.DataBindings.Add("Text", bindingSource1, "CCAPV");
            textBox36.DataBindings.Add("Text", bindingSource1, "NOTES");

            // Full Charge-6 Bindings ("T1Mode, T1Time1Hr, T1Time1Min, T1Curr1, T1Volts1, T1Time2Hr, T1Time2Min, T1Curr2, T1Volts2, T1Ohms")
            comboBox2.DataBindings.Add("Text", bindingSource1, "T1Mode");
            numericUpDown1.DataBindings.Add("Text", bindingSource1, "T1Time1Hr");
            numericUpDown2.DataBindings.Add("Text", bindingSource1, "T1Time1Min");
            numericUpDown3.DataBindings.Add("Text", bindingSource1, "T1Curr1");
            numericUpDown4.DataBindings.Add("Text", bindingSource1, "T1Volts1");
            numericUpDown8.DataBindings.Add("Text", bindingSource1, "T1Time2Hr");
            numericUpDown7.DataBindings.Add("Text", bindingSource1, "T1Time2Min");
            numericUpDown6.DataBindings.Add("Text", bindingSource1, "T1Curr2");
            numericUpDown5.DataBindings.Add("Text", bindingSource1, "T1Volts2");
            // Full Charge-4 Bindings ("T2Mode, T2Time1Hr, T2Time1Min, T2Curr1, T2Volts1, T2Time2Hr, T2Time2Min, T2Curr2, T2Volts2, T2Ohms")
            comboBox1.DataBindings.Add("Text", bindingSource1, "T2Mode");
            numericUpDown16.DataBindings.Add("Text", bindingSource1, "T2Time1Hr");
            numericUpDown15.DataBindings.Add("Text", bindingSource1, "T2Time1Min");
            numericUpDown14.DataBindings.Add("Text", bindingSource1, "T2Curr1");
            numericUpDown13.DataBindings.Add("Text", bindingSource1, "T2Volts1");
            numericUpDown12.DataBindings.Add("Text", bindingSource1, "T2Time2Hr");
            numericUpDown11.DataBindings.Add("Text", bindingSource1, "T2Time2Min");
            numericUpDown10.DataBindings.Add("Text", bindingSource1, "T2Curr2");
            numericUpDown9.DataBindings.Add("Text", bindingSource1, "T2Volts2");
            // Top Charge-4 Bindings ("T3Mode, T3Time1Hr, T3Time1Min, T3Curr1, T3Volts1, T3Time2Hr, T3Time2Min, T3Curr2, T3Volts2, T3Ohms")
            comboBox3.DataBindings.Add("Text", bindingSource1, "T3Mode");
            //numericUpDown24.DataBindings.Add("Text", bindingSource1, "T3Time1Hr");
            numericUpDown23.DataBindings.Add("Text", bindingSource1, "T3Time1Min");
            numericUpDown22.DataBindings.Add("Text", bindingSource1, "T3Curr1");
            numericUpDown21.DataBindings.Add("Text", bindingSource1, "T3Volts1");
            numericUpDown20.DataBindings.Add("Text", bindingSource1, "T3Time2Hr");
            numericUpDown19.DataBindings.Add("Text", bindingSource1, "T3Time2Min");
            numericUpDown18.DataBindings.Add("Text", bindingSource1, "T3Curr2");
            numericUpDown17.DataBindings.Add("Text", bindingSource1, "T3Volts2");
            // Top Charge-2 Bindings ("T4Mode, T4Time1Hr, T4Time1Min, T4Curr1, T4Volts1, T4Time2Hr, T4Time2Min, T4Curr2, T4Volts2, T4Ohms")
            comboBox4.DataBindings.Add("Text", bindingSource1, "T4Mode");
            //numericUpDown32.DataBindings.Add("Text", bindingSource1, "T4Time1Hr");
            numericUpDown31.DataBindings.Add("Text", bindingSource1, "T4Time1Min");
            numericUpDown30.DataBindings.Add("Text", bindingSource1, "T4Curr1");
            numericUpDown29.DataBindings.Add("Text", bindingSource1, "T4Volts1");
            numericUpDown28.DataBindings.Add("Text", bindingSource1, "T4Time2Hr");
            numericUpDown27.DataBindings.Add("Text", bindingSource1, "T4Time2Min");
            numericUpDown26.DataBindings.Add("Text", bindingSource1, "T4Curr2");
            numericUpDown25.DataBindings.Add("Text", bindingSource1, "T4Volts2");
            // Top Charge-1 Bindings ("T5Mode, T5Time1Hr, T5Time1Min, T5Curr1, T5Volts1, T5Time2Hr, T5Time2Min, T5Curr2, T5Volts2, T5Ohms")
            comboBox5.DataBindings.Add("Text", bindingSource1, "T5Mode");
            //numericUpDown40.DataBindings.Add("Text", bindingSource1, "T5Time1Hr");
            numericUpDown39.DataBindings.Add("Text", bindingSource1, "T5Time1Min");
            numericUpDown38.DataBindings.Add("Text", bindingSource1, "T5Curr1");
            numericUpDown37.DataBindings.Add("Text", bindingSource1, "T5Volts1");
            numericUpDown36.DataBindings.Add("Text", bindingSource1, "T5Time2Hr");
            numericUpDown35.DataBindings.Add("Text", bindingSource1, "T5Time2Min");
            numericUpDown34.DataBindings.Add("Text", bindingSource1, "T5Curr2");
            numericUpDown33.DataBindings.Add("Text", bindingSource1, "T5Volts2");
            // Capacity-1 Bindings ("T6Mode, T6Time1Hr, T6Time1Min, T6Curr1, T6Volts1, T6Time2Hr, T6Time2Min, T6Curr2, T6Volts2, T6Ohms")
            comboBox6.DataBindings.Add("Text", bindingSource1, "T6Mode");
            //numericUpDown45.DataBindings.Add("Text", bindingSource1, "T6Time1Hr");
            numericUpDown44.DataBindings.Add("Text", bindingSource1, "T6Time1Min");
            numericUpDown43.DataBindings.Add("Text", bindingSource1, "T6Curr1");
            numericUpDown42.DataBindings.Add("Text", bindingSource1, "T6Volts1");
            numericUpDown41.DataBindings.Add("Text", bindingSource1, "T6Ohms");
            // Discharge Bindings ("T7Mode, T7Time1Hr, T7Time1Min, T7Curr1, T7Volts1, T7Time2Hr, T7Time2Min, T7Curr2, T7Volts2, T7Ohms")
            //comboBox7.DataBindings.Add("Text", bindingSource1, "T7Mode");
            //numericUpDown50.DataBindings.Add("Text", bindingSource1, "T7Time1Hr");
            numericUpDown49.DataBindings.Add("Text", bindingSource1, "T7Time1Min");
            numericUpDown48.DataBindings.Add("Text", bindingSource1, "T7Curr1");
            numericUpDown47.DataBindings.Add("Text", bindingSource1, "T7Volts1");
            numericUpDown46.DataBindings.Add("Text", bindingSource1, "T7Ohms");
            // Slow Charge-14 Bindings ("T8Mode, T8Time1Hr, T8Time1Min, T8Curr1, T8Volts1, T8Time2Hr, T8Time2Min, T8Curr2, T8Volts2, T8Ohms")
            comboBox9.DataBindings.Add("Text", bindingSource1, "T8Mode");
            //numericUpDown63.DataBindings.Add("Text", bindingSource1, "T8Time1Hr");
            numericUpDown62.DataBindings.Add("Text", bindingSource1, "T8Time1Min");
            numericUpDown61.DataBindings.Add("Text", bindingSource1, "T8Curr1");
            numericUpDown60.DataBindings.Add("Text", bindingSource1, "T8Volts1");
            numericUpDown59.DataBindings.Add("Text", bindingSource1, "T8Time2Hr");
            numericUpDown58.DataBindings.Add("Text", bindingSource1, "T8Time2Min");
            numericUpDown57.DataBindings.Add("Text", bindingSource1, "T8Curr2");
            numericUpDown56.DataBindings.Add("Text", bindingSource1, "T8Volts2");
            // Slow Charge-16 Bindings ("T9Mode, T9Time1Hr, T9Time1Min, T9Curr1, T9Volts1, T9Time2Hr, T9Time2Min, T9Curr2, T9Volts2, T9Ohms")
            comboBox10.DataBindings.Add("Text", bindingSource1, "T9Mode");
            //numericUpDown71.DataBindings.Add("Text", bindingSource1, "T9Time1Hr");
            numericUpDown70.DataBindings.Add("Text", bindingSource1, "T9Time1Min");
            numericUpDown69.DataBindings.Add("Text", bindingSource1, "T9Curr1");
            numericUpDown68.DataBindings.Add("Text", bindingSource1, "T9Volts1");
            numericUpDown67.DataBindings.Add("Text", bindingSource1, "T9Time2Hr");
            numericUpDown66.DataBindings.Add("Text", bindingSource1, "T9Time2Min");
            numericUpDown65.DataBindings.Add("Text", bindingSource1, "T9Curr2");
            numericUpDown64.DataBindings.Add("Text", bindingSource1, "T9Volts2");
            // Custom Charge Bindings ("T10Mode, T10Time1Hr, T10Time1Min, T10Curr1, T10Volts1, T10Time2Hr, T10Time2Min, T10Curr2, T10Volts2, T10Ohms")
            comboBox11.DataBindings.Add("Text", bindingSource1, "T10Mode");
            numericUpDown79.DataBindings.Add("Text", bindingSource1, "T10Time1Hr");
            numericUpDown78.DataBindings.Add("Text", bindingSource1, "T10Time1Min");
            numericUpDown77.DataBindings.Add("Text", bindingSource1, "T10Curr1");
            numericUpDown76.DataBindings.Add("Text", bindingSource1, "T10Volts1");
            numericUpDown75.DataBindings.Add("Text", bindingSource1, "T10Time2Hr");
            numericUpDown74.DataBindings.Add("Text", bindingSource1, "T10Time2Min");
            numericUpDown73.DataBindings.Add("Text", bindingSource1, "T10Curr2");
            numericUpDown72.DataBindings.Add("Text", bindingSource1, "T10Volts2");
            // Custom Cap Bindings ("T11Mode, T11Time1Hr, T11Time1Min, T11Curr1, T11Volts1, T11Time2Hr, T11Time2Min, T11Curr2, T11Volts2, T11Ohms")
            comboBox8.DataBindings.Add("Text", bindingSource1, "T11Mode");
            numericUpDown55.DataBindings.Add("Text", bindingSource1, "T11Time1Hr");
            numericUpDown54.DataBindings.Add("Text", bindingSource1, "T11Time1Min");
            numericUpDown53.DataBindings.Add("Text", bindingSource1, "T11Curr1");
            numericUpDown52.DataBindings.Add("Text", bindingSource1, "T11Volts1");
            numericUpDown51.DataBindings.Add("Text", bindingSource1, "T11Ohms");
            // Constant Voltage Bindings ("T12Mode, T12Time1Hr, T12Time1Min, T12Curr1, T12Volts1, T12Time2Hr, T12Time2Min, T12Curr2, T12Volts2, T12Ohms")
            //comboBox12.DataBindings.Add("Text", bindingSource1, "T12Mode");
            //numericUpDown87.DataBindings.Add("Text", bindingSource1, "T12Time1Hr");
            numericUpDown86.DataBindings.Add("Text", bindingSource1, "T12Time1Min");
            numericUpDown85.DataBindings.Add("Text", bindingSource1, "T12Curr1");
            numericUpDown84.DataBindings.Add("Text", bindingSource1, "T12Volts1");
            numericUpDown83.DataBindings.Add("Text", bindingSource1, "T12Time2Hr");
            numericUpDown82.DataBindings.Add("Text", bindingSource1, "T12Time2Min");
            numericUpDown81.DataBindings.Add("Text", bindingSource1, "T12Curr2");
            numericUpDown80.DataBindings.Add("Text", bindingSource1, "T12Volts2");
            // Full Charge-4.5 Bindings ("T13Mode, T13Time1Hr, T13Time1Min, T13Curr1, T13Volts1, T13Time2Hr, T13Time2Min, T13Curr2, T13Volts2, T13Ohms")
            //comboBox14.DataBindings.Add("Text", bindingSource1, "T13Mode");
            //numericUpDown98.DataBindings.Add("Text", bindingSource1, "T13Time1Hr");
            numericUpDown97.DataBindings.Add("Text", bindingSource1, "T13Time1Min");
            numericUpDown96.DataBindings.Add("Text", bindingSource1, "T13Curr1");
            numericUpDown95.DataBindings.Add("Text", bindingSource1, "T13Volts1");
            //numericUpDown94.DataBindings.Add("Text", bindingSource1, "T13Time2Hr");
            numericUpDown93.DataBindings.Add("Text", bindingSource1, "T13Time2Min");
            numericUpDown92.DataBindings.Add("Text", bindingSource1, "T13Curr2");
            numericUpDown91.DataBindings.Add("Text", bindingSource1, "T13Volts2");
            // Custom Charge 2 Bindings ("T14Mode, T14Time1Hr, T14Time1Min, T14Curr1, T14Volts1, T14Time2Hr, T14Time2Min, T14Curr2, T14Volts2, T14Ohms")
            comboBox15.DataBindings.Add("Text", bindingSource1, "T14Mode");
            numericUpDown106.DataBindings.Add("Text", bindingSource1, "T14Time1Hr");
            numericUpDown105.DataBindings.Add("Text", bindingSource1, "T14Time1Min");
            numericUpDown104.DataBindings.Add("Text", bindingSource1, "T14Curr1");
            numericUpDown103.DataBindings.Add("Text", bindingSource1, "T14Volts1");
            numericUpDown102.DataBindings.Add("Text", bindingSource1, "T14Time2Hr");
            numericUpDown101.DataBindings.Add("Text", bindingSource1, "T14Time2Min");
            numericUpDown100.DataBindings.Add("Text", bindingSource1, "T14Curr2");
            numericUpDown99.DataBindings.Add("Text", bindingSource1, "T14Volts2");
            // Custom Charge 3 Bindings ("T15Mode, T15Time1Hr, T15Time1Min, T15Curr1, T15Volts1, T15Time2Hr, T15Time2Min, T15Curr2, T15Volts2, T15Ohms")
            comboBox16.DataBindings.Add("Text", bindingSource1, "T15Mode");
            numericUpDown114.DataBindings.Add("Text", bindingSource1, "T15Time1Hr");
            numericUpDown113.DataBindings.Add("Text", bindingSource1, "T15Time1Min");
            numericUpDown112.DataBindings.Add("Text", bindingSource1, "T15Curr1");
            numericUpDown111.DataBindings.Add("Text", bindingSource1, "T15Volts1");
            numericUpDown110.DataBindings.Add("Text", bindingSource1, "T15Time2Hr");
            numericUpDown109.DataBindings.Add("Text", bindingSource1, "T15Time2Min");
            numericUpDown108.DataBindings.Add("Text", bindingSource1, "T15Curr2");
            numericUpDown107.DataBindings.Add("Text", bindingSource1, "T15Volts2");

            

            #endregion

            #region setup the combo box
            ComboBox CustomerCB = toolStripCBBats.ComboBox;
            CustomerCB.DisplayMember = "BatteryModel";
            CustomerCB.DataSource = bindingSource1;


            #endregion

        }

        private void SetupForms()
        {
            // set up the numeric up down bounds
            // Charge Time 1
            numericUpDown1.Minimum = 0;             //hours
            numericUpDown1.Maximum = 2;
            numericUpDown2.Minimum = 0;             //mins
            numericUpDown2.Maximum = 59;
            numericUpDown3.Minimum = 0;             //charge current 1
            numericUpDown3.Maximum = 50;
            numericUpDown4.Minimum = 0;             //charge voltage 1
            numericUpDown4.Maximum = 77;

            // Charge Time 2
            numericUpDown8.Minimum = 0;             //hours
            numericUpDown8.Maximum = 6;
            numericUpDown7.Minimum = 0;             //mins
            numericUpDown7.Maximum = 59;
            numericUpDown6.Minimum = 0;             //charge current 2
            numericUpDown6.Maximum = 50;
            numericUpDown5.Minimum = 0;             //charge voltage 2
            numericUpDown5.Maximum = 77;

            //fixed controls
            comboBox12.Text = "12 Constant Voltage";
            comboBox7.Text = "30 Full Discharge";


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
            updateCurVals();
        }

        // we use this bool to allow us to allow the databinding indext to be changed...
        bool Inhibit = true;
        bool InhibitCB = true;

        private void toolStripCBCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (InhibitCB) { return; }

            //Validate before moving
            if (ValidateIt())
            {
                InhibitCB = true;
                // move back
                toolStripCBBats.SelectedIndex = bindingNavigator1.BindingSource.Position;
                InhibitCB = false;

            }
            else
            {
                InhibitCB = false;
                updateCurVals();
            }
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && bindingNavigator1.BindingSource.Position < CustomBats.Tables[0].Rows.Count)
                {
                    CustomBats.Tables[0].Rows[CustomBats.Tables[0].Rows.Count - 1].Delete();
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

            //we need to make sure all of the tabs have been "show"n first
            //this is because the binding source doesn't update until the tab has been selected, which was killing saved values!
            int selected = tabControl1.SelectedIndex;
            LockWindowUpdate(this.Handle);
            foreach (TabPage tp in tabControl1.TabPages)
            {
                tp.Show();
            }


            tabControl1.SelectTab(0);
            this.BeginInvoke(new Action(() =>
            {
                tabControl1.SelectTab(selected);
                LockWindowUpdate(IntPtr.Zero);
            }));

            // update the current vars....
            //current data..
            curTemp1 = textBox2.Text;
            curTemp2 = textBox1.Text;
            curTemp3 = textBox3.Text;
            curTemp4 = comboBox13.Text;

            curTemp5 = numericUpDown89.Text;
            curTemp6 = numericUpDown88.Text;
            curTemp7 = numericUpDown90.Text;

            curTemp10 = textBox27.Text;
            curTemp11 = textBox28.Text;
            curTemp12 = textBox29.Text;

            curTemp13 = textBox21.Text;
            curTemp14 = textBox22.Text;
            curTemp15 = textBox23.Text;

            curTemp16 = textBox36.Text;

            curTemp17 = comboBox2.Text;

            curTemp18 = numericUpDown1.Value;
            curTemp19 = numericUpDown2.Value;
            curTemp20 = numericUpDown3.Value;
            curTemp21 = numericUpDown4.Value;

            curTemp22 = numericUpDown8.Value;
            curTemp23 = numericUpDown7.Value;
            curTemp24 = numericUpDown6.Value;
            curTemp25 = numericUpDown5.Value;

            curTemp26 = comboBox1.Text;

            curTemp27 = numericUpDown16.Value;
            curTemp28 = numericUpDown15.Value;
            curTemp29 = numericUpDown14.Value;
            curTemp30 = numericUpDown13.Value;

            curTemp31 = numericUpDown12.Value;
            curTemp32 = numericUpDown11.Value;
            curTemp33 = numericUpDown10.Value;
            curTemp34 = numericUpDown9.Value;

            curTemp35 = comboBox3.Text;

            curTemp36 = numericUpDown24.Value;
            curTemp37 = numericUpDown23.Value;
            curTemp38 = numericUpDown22.Value;
            curTemp39 = numericUpDown21.Value;

            curTemp40 = numericUpDown20.Value;
            curTemp41 = numericUpDown19.Value;
            curTemp42 = numericUpDown18.Value;
            curTemp43 = numericUpDown17.Value;

            curTemp44 = comboBox4.Text;

            curTemp45 = numericUpDown32.Value;
            curTemp46 = numericUpDown31.Value;
            curTemp47 = numericUpDown30.Value;
            curTemp48 = numericUpDown29.Value;

            curTemp49 = numericUpDown28.Value;
            curTemp50 = numericUpDown27.Value;
            curTemp51 = numericUpDown26.Value;
            curTemp52 = numericUpDown25.Value;

            curTemp53 = comboBox5.Text;

            curTemp54 = numericUpDown40.Value;
            curTemp55 = numericUpDown39.Value;
            curTemp56 = numericUpDown38.Value;
            curTemp57 = numericUpDown37.Value;

            curTemp58 = numericUpDown36.Value;
            curTemp59 = numericUpDown35.Value;
            curTemp60 = numericUpDown34.Value;
            curTemp61 = numericUpDown33.Value;

            curTemp62 = comboBox12.Text;

            curTemp63 = numericUpDown87.Value;
            curTemp64 = numericUpDown86.Value;
            curTemp65 = numericUpDown85.Value;
            curTemp66 = numericUpDown84.Value;

            curTemp67 = numericUpDown83.Value;
            curTemp68 = numericUpDown82.Value;
            curTemp69 = numericUpDown81.Value;
            curTemp70 = numericUpDown80.Value;

            curTemp71 = comboBox6.Text;

            curTemp72 = numericUpDown45.Value;
            curTemp73 = numericUpDown44.Value;
            curTemp74 = numericUpDown43.Value;
            curTemp75 = numericUpDown42.Value;
            curTemp76 = numericUpDown41.Value;

            curTemp77 = comboBox7.Text;

            curTemp78 = numericUpDown50.Value;
            curTemp79 = numericUpDown49.Value;
            curTemp80 = numericUpDown48.Value;
            curTemp81 = numericUpDown47.Value;
            curTemp82 = numericUpDown46.Value;

            curTemp83 = comboBox9.Text;

            curTemp84 = numericUpDown63.Value;
            curTemp85 = numericUpDown62.Value;
            curTemp86 = numericUpDown61.Value;
            curTemp87 = numericUpDown60.Value;

            curTemp88 = numericUpDown59.Value;
            curTemp89 = numericUpDown58.Value;
            curTemp90 = numericUpDown57.Value;
            curTemp101 = numericUpDown56.Value;

            curTemp91 = comboBox10.Text;

            curTemp92 = numericUpDown71.Value;
            curTemp93 = numericUpDown70.Value;
            curTemp94 = numericUpDown69.Value;
            curTemp95 = numericUpDown68.Value;

            curTemp96 = numericUpDown67.Value;
            curTemp97 = numericUpDown66.Value;
            curTemp98 = numericUpDown65.Value;
            curTemp99 = numericUpDown64.Value;

            curTemp100 = comboBox11.Text;

            curTemp102 = numericUpDown79.Value;
            curTemp103 = numericUpDown78.Value;
            curTemp104 = numericUpDown77.Value;
            curTemp105 = numericUpDown76.Value;

            curTemp106 = numericUpDown75.Value;
            curTemp107 = numericUpDown74.Value;
            curTemp108 = numericUpDown73.Value;
            curTemp109 = numericUpDown72.Value;

            curTemp110 = comboBox8.Text;

            curTemp111 = numericUpDown55.Value;
            curTemp112 = numericUpDown54.Value;
            curTemp113 = numericUpDown53.Value;
            curTemp114 = numericUpDown52.Value;
            curTemp115 = numericUpDown51.Value;

            curTemp116 = comboBox14.Text;

            curTemp117 = numericUpDown96.Value;
            curTemp118 = numericUpDown95.Value;
            curTemp119 = numericUpDown92.Value;
            curTemp120 = numericUpDown91.Value;

            curTemp121 = comboBox15.Text;

            curTemp122 = numericUpDown106.Value;
            curTemp123 = numericUpDown105.Value;
            curTemp124 = numericUpDown104.Value;
            curTemp125 = numericUpDown103.Value;

            curTemp126 = numericUpDown102.Value;
            curTemp127 = numericUpDown101.Value;
            curTemp128 = numericUpDown100.Value;
            curTemp129 = numericUpDown99.Value;

            curTemp130 = comboBox16.Text;

            curTemp131 = numericUpDown114.Value;
            curTemp132 = numericUpDown113.Value;
            curTemp133 = numericUpDown112.Value;
            curTemp134 = numericUpDown111.Value;

            curTemp135 = numericUpDown110.Value;
            curTemp136 = numericUpDown109.Value;
            curTemp137 = numericUpDown108.Value;
            curTemp138 = numericUpDown107.Value;
        }
        

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

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
                if (MessageBox.Show(this, "Are you sure you want to remove this battery from the data base?", "Delete Record", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    // set up the db Connection
                    string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                    OleDbConnection conn = new OleDbConnection(connectionString);


                    //get the current row
                    DataRowView current = (DataRowView)bindingSource1.Current;

                    // first test to see if the record already is actually in the database

                    if (current["RecordID"].ToString() != "")
                    {
                        //record already exist as we need to do an update

                        string cmdStr = "DELETE FROM BatteriesCustom WHERE RecordID=" + current["RecordID"].ToString();
                        OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                        lock (Main_Form.dataBaseLock)
                        {
                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }

                        // Also update the binding source
                        CustomBats.Tables[0].Rows[bindingNavigator1.BindingSource.Position].Delete();

                    }
                    else
                    {
                        MessageBox.Show(this, "That record was not in the DB. You must save it in order to delete it.","Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Deletion Error" + ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                updateCurVals();
            }
        }

        private void saveToolStripButton_Click_1(object sender, EventArgs e)
        {
            Inhibit = true;

            //to get around the new entry issue...
            #region empty database work around...
            if (bindingNavigator1.BindingSource.Position == -1)
            {
                string temp1 = textBox2.Text;
                string temp2 = textBox1.Text;
                string temp3 = textBox3.Text;
                string temp4 = comboBox13.Text;

                string temp5 = numericUpDown89.Text;
                string temp6 = numericUpDown88.Text;
                string temp7 = numericUpDown90.Text;

                string temp10 = textBox27.Text;
                string temp11 = textBox28.Text;
                string temp12 = textBox29.Text;

                string temp13 = textBox21.Text;
                string temp14 = textBox22.Text;
                string temp15 = textBox23.Text;

                string temp16 = textBox36.Text;

                string temp17 = comboBox2.Text;

                decimal temp18 = numericUpDown1.Value;
                decimal temp19 = numericUpDown2.Value;
                decimal temp20 = numericUpDown3.Value;
                decimal temp21 = numericUpDown4.Value;

                decimal temp22 = numericUpDown8.Value;
                decimal temp23 = numericUpDown7.Value;
                decimal temp24 = numericUpDown6.Value;
                decimal temp25 = numericUpDown5.Value;

                string temp26 = comboBox1.Text;

                decimal temp27 = numericUpDown16.Value;
                decimal temp28 = numericUpDown15.Value;
                decimal temp29 = numericUpDown14.Value;
                decimal temp30 = numericUpDown13.Value;

                decimal temp31 = numericUpDown12.Value;
                decimal temp32 = numericUpDown11.Value;
                decimal temp33 = numericUpDown10.Value;
                decimal temp34 = numericUpDown9.Value;

                string temp35 = comboBox3.Text;

                decimal temp36 = numericUpDown24.Value;
                decimal temp37 = numericUpDown23.Value;
                decimal temp38 = numericUpDown22.Value;
                decimal temp39 = numericUpDown21.Value;

                decimal temp40 = numericUpDown20.Value;
                decimal temp41 = numericUpDown19.Value;
                decimal temp42 = numericUpDown18.Value;
                decimal temp43 = numericUpDown17.Value;

                string temp44 = comboBox4.Text;

                decimal temp45 = numericUpDown32.Value;
                decimal temp46 = numericUpDown31.Value;
                decimal temp47 = numericUpDown30.Value;
                decimal temp48 = numericUpDown29.Value;

                decimal temp49 = numericUpDown28.Value;
                decimal temp50 = numericUpDown27.Value;
                decimal temp51 = numericUpDown26.Value;
                decimal temp52 = numericUpDown25.Value;

                string temp53 = comboBox5.Text;

                decimal temp54 = numericUpDown40.Value;
                decimal temp55 = numericUpDown39.Value;
                decimal temp56 = numericUpDown38.Value;
                decimal temp57 = numericUpDown37.Value;

                decimal temp58 = numericUpDown36.Value;
                decimal temp59 = numericUpDown35.Value;
                decimal temp60 = numericUpDown34.Value;
                decimal temp61 = numericUpDown33.Value;

                string temp62 = comboBox12.Text;

                decimal temp63 = numericUpDown87.Value;
                decimal temp64 = numericUpDown86.Value;
                decimal temp65 = numericUpDown85.Value;
                decimal temp66 = numericUpDown84.Value;

                decimal temp67 = numericUpDown83.Value;
                decimal temp68 = numericUpDown82.Value;
                decimal temp69 = numericUpDown81.Value;
                decimal temp70 = numericUpDown80.Value;

                string temp71 = comboBox6.Text;

                decimal temp72 = numericUpDown45.Value;
                decimal temp73 = numericUpDown44.Value;
                decimal temp74 = numericUpDown43.Value;
                decimal temp75 = numericUpDown42.Value;
                decimal temp76 = numericUpDown41.Value;

                string temp77 = comboBox7.Text;

                decimal temp78 = numericUpDown50.Value;
                decimal temp79 = numericUpDown49.Value;
                decimal temp80 = numericUpDown48.Value;
                decimal temp81 = numericUpDown47.Value;
                decimal temp82 = numericUpDown46.Value;

                string temp83 = comboBox9.Text;

                decimal temp84 = numericUpDown63.Value;
                decimal temp85 = numericUpDown62.Value;
                decimal temp86 = numericUpDown61.Value;
                decimal temp87 = numericUpDown60.Value;

                decimal temp88 = numericUpDown59.Value;
                decimal temp89 = numericUpDown58.Value;
                decimal temp90 = numericUpDown57.Value;
                decimal temp101 = numericUpDown56.Value;

                string temp91 = comboBox10.Text;

                decimal temp92 = numericUpDown71.Value;
                decimal temp93 = numericUpDown70.Value;
                decimal temp94 = numericUpDown69.Value;
                decimal temp95 = numericUpDown68.Value;

                decimal temp96 = numericUpDown67.Value;
                decimal temp97 = numericUpDown66.Value;
                decimal temp98 = numericUpDown65.Value;
                decimal temp99 = numericUpDown64.Value;

                string temp100 = comboBox11.Text;

                decimal temp102 = numericUpDown79.Value;
                decimal temp103 = numericUpDown78.Value;
                decimal temp104 = numericUpDown77.Value;
                decimal temp105 = numericUpDown76.Value;

                decimal temp106 = numericUpDown75.Value;
                decimal temp107 = numericUpDown74.Value;
                decimal temp108 = numericUpDown73.Value;
                decimal temp109 = numericUpDown72.Value;

                string temp110 = comboBox8.Text;

                decimal temp111 = numericUpDown55.Value;
                decimal temp112 = numericUpDown54.Value;
                decimal temp113 = numericUpDown53.Value;
                decimal temp114 = numericUpDown52.Value;
                decimal temp115 = numericUpDown51.Value;

                string temp116 = comboBox14.Text;

                decimal temp117 = numericUpDown96.Value;
                decimal temp118 = numericUpDown95.Value;
                decimal temp119 = numericUpDown92.Value;
                decimal temp120 = numericUpDown91.Value;

                string temp121 = comboBox15.Text;

                decimal temp122 = numericUpDown106.Value;
                decimal temp123 = numericUpDown105.Value;
                decimal temp124 = numericUpDown104.Value;
                decimal temp125 = numericUpDown103.Value;

                decimal temp126 = numericUpDown102.Value;
                decimal temp127 = numericUpDown101.Value;
                decimal temp128 = numericUpDown100.Value;
                decimal temp129 = numericUpDown99.Value;

                string temp130 = comboBox16.Text;

                decimal temp131 = numericUpDown114.Value;
                decimal temp132 = numericUpDown113.Value;
                decimal temp133 = numericUpDown112.Value;
                decimal temp134 = numericUpDown111.Value;

                decimal temp135 = numericUpDown110.Value;
                decimal temp136 = numericUpDown109.Value;
                decimal temp137 = numericUpDown108.Value;
                decimal temp138 = numericUpDown107.Value;
                
                bindingNavigator1.BindingSource.AddNew();
                bindingNavigator1.BindingSource.Position = 0;

                textBox2.Text = temp1;
                textBox1.Text = temp2;
                textBox3.Text = temp3;
                comboBox13.Text = temp4;

                numericUpDown89.Text = temp5;
                numericUpDown88.Text = temp6;
                numericUpDown90.Text = temp7;

                textBox27.Text = temp10;
                textBox28.Text = temp11;
                textBox29.Text = temp12;

                textBox21.Text = temp13;
                textBox22.Text = temp14;
                textBox23.Text = temp15;

                textBox36.Text = temp16;

                comboBox2.Text = temp17;

                numericUpDown1.Value = temp18;
                numericUpDown2.Value = temp19;
                numericUpDown3.Value = temp20;
                numericUpDown4.Value = temp21;

                numericUpDown8.Value = temp22;
                numericUpDown7.Value = temp23;
                numericUpDown6.Value = temp24;
                numericUpDown5.Value = temp25;

                comboBox1.Text = temp26;

                numericUpDown16.Value = temp27;
                numericUpDown15.Value = temp28;
                numericUpDown14.Value = temp29;
                numericUpDown13.Value = temp30;

                numericUpDown12.Value = temp31;
                numericUpDown11.Value = temp32;
                numericUpDown10.Value = temp33;
                numericUpDown9.Value = temp34;

                comboBox3.Text = temp35;

                numericUpDown24.Value = temp36;
                numericUpDown23.Value = temp37;
                numericUpDown22.Value = temp38;
                numericUpDown21.Value = temp39;

                numericUpDown20.Value = temp40;
                numericUpDown19.Value = temp41;
                numericUpDown18.Value = temp42;
                numericUpDown17.Value = temp43;

                comboBox4.Text = temp44;

                numericUpDown32.Value = temp45;
                numericUpDown31.Value = temp46;
                numericUpDown30.Value = temp47;
                numericUpDown29.Value = temp48;

                numericUpDown28.Value = temp49;
                numericUpDown27.Value = temp50;
                numericUpDown26.Value = temp51;
                numericUpDown25.Value = temp52;

                comboBox5.Text = temp53;

                numericUpDown40.Value = temp54;
                numericUpDown39.Value = temp55;
                numericUpDown38.Value = temp56;
                numericUpDown37.Value = temp57;

                numericUpDown36.Value = temp58;
                numericUpDown35.Value = temp59;
                numericUpDown34.Value = temp60;
                numericUpDown33.Value = temp61;

                comboBox12.Text = temp62;

                numericUpDown87.Value = temp63;
                numericUpDown86.Value = temp64;
                numericUpDown85.Value = temp65;
                numericUpDown84.Value = temp66;

                numericUpDown83.Value = temp67;
                numericUpDown82.Value = temp68;
                numericUpDown81.Value = temp69;
                numericUpDown80.Value = temp70;

                comboBox6.Text = temp71;

                numericUpDown45.Value = temp72;
                numericUpDown44.Value = temp73;
                numericUpDown43.Value = temp74;
                numericUpDown42.Value = temp75;
                numericUpDown41.Value = temp76;

                comboBox7.Text = temp77;

                numericUpDown50.Value = temp78;
                numericUpDown49.Value = temp79;
                numericUpDown48.Value = temp80;
                numericUpDown47.Value = temp81;
                numericUpDown46.Value = temp82;

                comboBox9.Text = temp83;

                numericUpDown63.Value = temp84;
                numericUpDown62.Value = temp85;
                numericUpDown61.Value = temp86;
                numericUpDown60.Value = temp87;

                numericUpDown59.Value = temp88;
                numericUpDown58.Value = temp89;
                numericUpDown57.Value = temp90;
                numericUpDown56.Value = temp101;

                comboBox10.Text = temp91;

                numericUpDown71.Value = temp92;
                numericUpDown70.Value = temp93;
                numericUpDown69.Value = temp94;
                numericUpDown68.Value = temp95;

                numericUpDown67.Value = temp96;
                numericUpDown66.Value = temp97;
                numericUpDown65.Value = temp98;
                numericUpDown64.Value = temp99;

                comboBox11.Text = temp100;

                numericUpDown79.Value = temp102;
                numericUpDown78.Value = temp103;
                numericUpDown77.Value = temp104;
                numericUpDown76.Value = temp105;

                numericUpDown75.Value = temp106;
                numericUpDown74.Value = temp107;
                numericUpDown73.Value = temp108;
                numericUpDown72.Value = temp109;

                comboBox8.Text = temp110;

                numericUpDown55.Value = temp111;
                numericUpDown54.Value = temp112;
                numericUpDown53.Value = temp113;
                numericUpDown52.Value = temp114;
                numericUpDown51.Value = temp115;

                comboBox15.Text = temp121;

                numericUpDown106.Value = temp122;
                numericUpDown105.Value = temp123;
                numericUpDown104.Value = temp124;
                numericUpDown103.Value = temp125;

                numericUpDown102.Value = temp126;
                numericUpDown101.Value = temp127;
                numericUpDown100.Value = temp128;
                numericUpDown99.Value = temp129;

                comboBox16.Text = temp130;

                numericUpDown114.Value = temp131;
                numericUpDown113.Value = temp132;
                numericUpDown112.Value = temp133;
                numericUpDown111.Value = temp134;

                numericUpDown110.Value = temp135;
                numericUpDown109.Value = temp136;
                numericUpDown108.Value = temp137;
                numericUpDown107.Value = temp138;

            }
            #endregion

            try
            {
                //we need to make sure all of the tabs have been "show"n first
                //this is because the binding source doesn't update until the tab has been selected, which was killing saved values!
                int selected = tabControl1.SelectedIndex;
                LockWindowUpdate(this.Handle);
                foreach (TabPage tp in tabControl1.TabPages)
                {
                    tp.Show();                    
                }

                
                tabControl1.SelectTab(0);
                this.BeginInvoke(new Action(() =>
                {
                    tabControl1.SelectTab(selected);
                    LockWindowUpdate(IntPtr.Zero);
                }));
                

                // set up the db Connection
                string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                OleDbConnection conn = new OleDbConnection(connectionString);

                //MAKE SURE YOU SELECT THE CURRENT ROW FOR DOUBLE SAVES!!!!!!!!!!!!!!!!!

                //get the current row
                DataRowView current = (DataRowView)bindingSource1.Current;

                // first test to see if the record already is in the database

                if (current["RecordID"].ToString() != "")
                {
                    //record already exist as we need to do an update

                    string cmdStr = "UPDATE BatteriesCustom SET BMFR='" + textBox1.Text.Replace("'", "''") +
                        "', BatteryModel='" + textBox2.Text.Replace("'", "''") +
                        "', BPN='" + textBox3.Text.Replace("'", "''") +
                        "', BTECH='" + comboBox13.Text.Replace("'", "''") +
                        "', VOLT='" + numericUpDown89.Text.Replace("'", "''") +
                        "', NCELLS='" + numericUpDown88.Text.Replace("'", "''") +
                        "', CAP='" + numericUpDown90.Text.Replace("'", "''") +
                        "', BCVMIN='" + textBox21.Text.Replace("'", "''") +
                        "', BCVMAX='" + textBox22.Text.Replace("'", "''") +
                        "', COT='" + textBox23.Text.Replace("'", "''") +
                        "', CCVMMIN='" + textBox27.Text.Replace("'", "''") +
                        "', CCVMAX='" + textBox28.Text.Replace("'", "''") +
                        "', CCAPV='" + textBox29.Text.Replace("'", "''") +
                        "', NOTES='" + textBox36.Text.Replace("'", "''") +
                        // Full Charge-6 ("T1Mode, T1Time1Hr, T1Time1Min, T1Curr1, T1Volts1, T1Time2Hr, T1Time2Min, T1Curr2, T1Volts2, T1Ohms")
                        "', T1Mode='" + comboBox2.Text.Replace("'", "''") +
                        "', T1Time1Hr='" + numericUpDown1.Text.Replace("'", "''") +
                        "', T1Time1Min='" + numericUpDown2.Text.Replace("'", "''") +
                        "', T1Curr1='" + numericUpDown3.Text.Replace("'", "''") +
                        "', T1Volts1='" + numericUpDown4.Text.Replace("'", "''") +
                        "', T1Time2Hr='" + numericUpDown8.Text.Replace("'", "''") +
                        "', T1Time2Min='" + numericUpDown7.Text.Replace("'", "''") +
                        "', T1Curr2='" + numericUpDown6.Text.Replace("'", "''") +
                        "', T1Volts2='" + numericUpDown5.Text.Replace("'", "''") +
                        // Full Charge-4 ("T2Mode, T2Time1Hr, T2Time1Min, T2Curr1, T2Volts1, T2Time2Hr, T2Time2Min, T2Curr2, T2Volts2, T2Ohms")
                        "', T2Mode='" + comboBox1.Text.Replace("'", "''") +
                        "', T2Time1Hr='" + numericUpDown16.Text.Replace("'", "''") +
                        "', T2Time1Min='" + numericUpDown15.Text.Replace("'", "''") +
                        "', T2Curr1='" + numericUpDown14.Text.Replace("'", "''") +
                        "', T2Volts1='" + numericUpDown13.Text.Replace("'", "''") +
                        "', T2Time2Hr='" + numericUpDown12.Text.Replace("'", "''") +
                        "', T2Time2Min='" + numericUpDown11.Text.Replace("'", "''") +
                        "', T2Curr2='" + numericUpDown10.Text.Replace("'", "''") +
                        "', T2Volts2='" + numericUpDown9.Text.Replace("'", "''") +
                        // Top Charge-4 ("T3Mode, T3Time1Hr, T3Time1Min, T3Curr1, T3Volts1, T3Time2Hr, T3Time2Min, T3Curr2, T3Volts2, T3Ohms")
                        "', T3Mode='" + comboBox3.Text.Replace("'", "''") +
                        "', T3Time1Hr='" + numericUpDown24.Text.Replace("'", "''") +
                        "', T3Time1Min='" + numericUpDown23.Text.Replace("'", "''") +
                        "', T3Curr1='" + numericUpDown22.Text.Replace("'", "''") +
                        "', T3Volts1='" + numericUpDown21.Text.Replace("'", "''") +
                        "', T3Time2Hr='" + numericUpDown20.Text.Replace("'", "''") +
                        "', T3Time2Min='" + numericUpDown19.Text.Replace("'", "''") +
                        "', T3Curr2='" + numericUpDown18.Text.Replace("'", "''") +
                        "', T3Volts2='" + numericUpDown17.Text.Replace("'", "''") +
                        // Top Charge-2 ("T4Mode, T4Time1Hr, T4Time1Min, T4Curr1, T4Volts1, T4Time2Hr, T4Time2Min, T4Curr2, T4Volts2, T4Ohms")
                        "', T4Mode='" + comboBox4.Text.Replace("'", "''") +
                        "', T4Time1Hr='" + numericUpDown32.Text.Replace("'", "''") +
                        "', T4Time1Min='" + numericUpDown31.Text.Replace("'", "''") +
                        "', T4Curr1='" + numericUpDown30.Text.Replace("'", "''") +
                        "', T4Volts1='" + numericUpDown29.Text.Replace("'", "''") +
                        "', T4Time2Hr='" + numericUpDown28.Text.Replace("'", "''") +
                        "', T4Time2Min='" + numericUpDown27.Text.Replace("'", "''") +
                        "', T4Curr2='" + numericUpDown26.Text.Replace("'", "''") +
                        "', T4Volts2='" + numericUpDown25.Text.Replace("'", "''") +
                        // Top Charge-1 ("T5Mode, T5Time1Hr, T5Time1Min, T5Curr1, T5Volts1, T5Time2Hr, T5Time2Min, T5Curr2, T5Volts2, T5Ohms")
                        "', T5Mode='" + comboBox5.Text.Replace("'", "''") +
                        "', T5Time1Hr='" + numericUpDown40.Text.Replace("'", "''") +
                        "', T5Time1Min='" + numericUpDown39.Text.Replace("'", "''") +
                        "', T5Curr1='" + numericUpDown38.Text.Replace("'", "''") +
                        "', T5Volts1='" + numericUpDown37.Text.Replace("'", "''") +
                        "', T5Time2Hr='" + numericUpDown36.Text.Replace("'", "''") +
                        "', T5Time2Min='" + numericUpDown35.Text.Replace("'", "''") +
                        "', T5Curr2='" + numericUpDown34.Text.Replace("'", "''") +
                        "', T5Volts2='" + numericUpDown33.Text.Replace("'", "''") +
                        // Capacity-1 ("T6Mode, T6Time1Hr, T6Time1Min, T6Curr1, T6Volts1, T6Time2Hr, T6Time2Min, T6Curr2, T6Volts2, T6Ohms")
                        "', T6Mode='" + comboBox6.Text.Replace("'", "''") +
                        "', T6Time1Hr='" + numericUpDown45.Text.Replace("'", "''") +
                        "', T6Time1Min='" + numericUpDown44.Text.Replace("'", "''") +
                        "', T6Curr1='" + numericUpDown43.Text.Replace("'", "''") +
                        "', T6Volts1='" + numericUpDown42.Text.Replace("'", "''") +
                        "', T6Ohms='" + numericUpDown41.Text.Replace("'", "''") +
                        // Discharge ("T7Mode, T7Time1Hr, T7Time1Min, T7Curr1, T7Volts1, T7Time2Hr, T7Time2Min, T7Curr2, T7Volts2, T7Ohms")
                        "', T7Mode='" + comboBox7.Text.Replace("'", "''") +
                        "', T7Time1Hr='" + numericUpDown50.Text.Replace("'", "''") +
                        "', T7Time1Min='" + numericUpDown49.Text.Replace("'", "''") +
                        "', T7Curr1='" + numericUpDown48.Text.Replace("'", "''") +
                        "', T7Volts1='" + numericUpDown47.Text.Replace("'", "''") +
                        "', T7Ohms='" + numericUpDown46.Text.Replace("'", "''") +
                        // Slow Charge-14 ("T8Mode, T8Time1Hr, T8Time1Min, T8Curr1, T8Volts1, T8Time2Hr, T8Time2Min, T8Curr2, T8Volts2, T8Ohms")
                        "', T8Mode='" + comboBox9.Text.Replace("'", "''") +
                        "', T8Time1Hr='" + numericUpDown63.Text.Replace("'", "''") +
                        "', T8Time1Min='" + numericUpDown62.Text.Replace("'", "''") +
                        "', T8Curr1='" + numericUpDown61.Text.Replace("'", "''") +
                        "', T8Volts1='" + numericUpDown60.Text.Replace("'", "''") +
                        "', T8Time2Hr='" + numericUpDown59.Text.Replace("'", "''") +
                        "', T8Time2Min='" + numericUpDown58.Text.Replace("'", "''") +
                        "', T8Curr2='" + numericUpDown57.Text.Replace("'", "''") +
                        "', T8Volts2='" + numericUpDown56.Text.Replace("'", "''") +
                        // Slow Charge-16 ("T9Mode, T9Time1Hr, T9Time1Min, T9Curr1, T9Volts1, T9Time2Hr, T9Time2Min, T9Curr2, T9Volts2, T9Ohms")
                        "', T9Mode='" + comboBox10.Text.Replace("'", "''") +
                        "', T9Time1Hr='" + numericUpDown71.Text.Replace("'", "''") +
                        "', T9Time1Min='" + numericUpDown70.Text.Replace("'", "''") +
                        "', T9Curr1='" + numericUpDown69.Text.Replace("'", "''") +
                        "', T9Volts1='" + numericUpDown68.Text.Replace("'", "''") +
                        "', T9Time2Hr='" + numericUpDown67.Text.Replace("'", "''") +
                        "', T9Time2Min='" + numericUpDown66.Text.Replace("'", "''") +
                        "', T9Curr2='" + numericUpDown65.Text.Replace("'", "''") +
                        "', T9Volts2='" + numericUpDown64.Text.Replace("'", "''") +
                        // Custom Charge ("T10Mode, T10Time1Hr, T10Time1Min, T10Curr1, T10Volts1, T10Time2Hr, T10Time2Min, T10Curr2, T10Volts2, T10Ohms")
                        "', T10Mode='" + comboBox11.Text.Replace("'", "''") +
                        "', T10Time1Hr='" + numericUpDown79.Text.Replace("'", "''") +
                        "', T10Time1Min='" + numericUpDown78.Text.Replace("'", "''") +
                        "', T10Curr1='" + numericUpDown77.Text.Replace("'", "''") +
                        "', T10Volts1='" + numericUpDown76.Text.Replace("'", "''") +
                        "', T10Time2Hr='" + numericUpDown75.Text.Replace("'", "''") +
                        "', T10Time2Min='" + numericUpDown74.Text.Replace("'", "''") +
                        "', T10Curr2='" + numericUpDown73.Text.Replace("'", "''") +
                        "', T10Volts2='" + numericUpDown72.Text.Replace("'", "''") +
                        // Custom Cap ("T11Mode, T11Time1Hr, T11Time1Min, T11Curr1, T11Volts1, T11Time2Hr, T11Time2Min, T11Curr2, T11Volts2, T11Ohms")
                        "', T11Mode='" + comboBox8.Text.Replace("'", "''") +
                        "', T11Time1Hr='" + numericUpDown55.Text.Replace("'", "''") +
                        "', T11Time1Min='" + numericUpDown54.Text.Replace("'", "''") +
                        "', T11Curr1='" + numericUpDown53.Text.Replace("'", "''") +
                        "', T11Volts1='" + numericUpDown52.Text.Replace("'", "''") +
                        "', T11Ohms='" + numericUpDown51.Text.Replace("'", "''") +
                        // Custom Charge ("T12Mode, T12Time1Hr, T12Time1Min, T12Curr1, T12Volts1, T12Time2Hr, T12Time2Min, T12Curr2, T12Volts2, T12Ohms")
                        "', T12Mode='" + comboBox12.Text.Replace("'", "''") +
                        "', T12Time1Hr='" + numericUpDown87.Text.Replace("'", "''") +
                        "', T12Time1Min='" + numericUpDown86.Text.Replace("'", "''") +
                        "', T12Curr1='" + numericUpDown85.Text.Replace("'", "''") +
                        "', T12Volts1='" + numericUpDown84.Text.Replace("'", "''") +
                        "', T12Time2Hr='" + numericUpDown83.Text.Replace("'", "''") +
                        "', T12Time2Min='" + numericUpDown82.Text.Replace("'", "''") +
                        "', T12Curr2='" + numericUpDown81.Text.Replace("'", "''") +
                        "', T12Volts2='" + numericUpDown80.Text.Replace("'", "''") +
                        // Full Charge-4.5 ("T13Mode, T13Time1Hr, T13Time1Min, T13Curr1, T13Volts1, T13Time2Hr, T13Time2Min, T13Curr2, T13Volts2, T13Ohms")
                        "', T13Mode='" + comboBox14.Text.Replace("'", "''") +
                        "', T13Time1Hr='" + numericUpDown98.Text.Replace("'", "''") +
                        "', T13Time1Min='" + numericUpDown97.Text.Replace("'", "''") +
                        "', T13Curr1='" + numericUpDown96.Text.Replace("'", "''") +
                        "', T13Volts1='" + numericUpDown95.Text.Replace("'", "''") +
                        "', T13Time2Hr='" + numericUpDown94.Text.Replace("'", "''") +
                        "', T13Time2Min='" + numericUpDown93.Text.Replace("'", "''") +
                        "', T13Curr2='" + numericUpDown92.Text.Replace("'", "''") +
                        "', T13Volts2='" + numericUpDown91.Text.Replace("'", "''") +
                        // finished with inputs!
                        "' WHERE RecordID=" + current["RecordID"].ToString();
                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                    cmdStr = "UPDATE BatteriesCustom SET" +
                        // Custom Charge 2("T14Mode, T14Time1Hr, T14Time1Min, T14Curr1, T14Volts1, T14Time2Hr, T14Time2Min, T14Curr2, T14Volts2, T14Ohms")
                        " T14Mode='" + comboBox15.Text.Replace("'", "''") +
                        "', T14Time1Hr='" + numericUpDown106.Text.Replace("'", "''") +
                        "', T14Time1Min='" + numericUpDown105.Text.Replace("'", "''") +
                        "', T14Curr1='" + numericUpDown104.Text.Replace("'", "''") +
                        "', T14Volts1='" + numericUpDown103.Text.Replace("'", "''") +
                        "', T14Time2Hr='" + numericUpDown102.Text.Replace("'", "''") +
                        "', T14Time2Min='" + numericUpDown101.Text.Replace("'", "''") +
                        "', T14Curr2='" + numericUpDown100.Text.Replace("'", "''") +
                        "', T14Volts2='" + numericUpDown99.Text.Replace("'", "''") +
                                            // Custom Charge 3("T15Mode, T15Time1Hr, T15Time1Min, T15Curr1, T15Volts1, T15Time2Hr, T15Time2Min, T15Curr2, T15Volts2, T15Ohms")
                        "', T15Mode='" + comboBox16.Text.Replace("'", "''") +
                        "', T15Time1Hr='" + numericUpDown114.Text.Replace("'", "''") +
                        "', T15Time1Min='" + numericUpDown113.Text.Replace("'", "''") +
                        "', T15Curr1='" + numericUpDown112.Text.Replace("'", "''") +
                        "', T15Volts1='" + numericUpDown111.Text.Replace("'", "''") +
                        "', T15Time2Hr='" + numericUpDown110.Text.Replace("'", "''") +
                        "', T15Time2Min='" + numericUpDown109.Text.Replace("'", "''") +
                        "', T15Curr2='" + numericUpDown108.Text.Replace("'", "''") +
                        "', T15Volts2='" + numericUpDown107.Text.Replace("'", "''") +
                                            // finished with inputs!
                        "' WHERE RecordID=" + current["RecordID"].ToString();
                    cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                    // Also update the model in the other tables!
                    cmdStr = "UPDATE Batteries SET BatteryModel='" + textBox2.Text.Replace("'", "''") + "' WHERE BatteryModel='" + current["BatteryModel"].ToString() + "'";
                    cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                    cmdStr = "UPDATE WorkOrders SET BatteryModel='" + textBox2.Text.Replace("'", "''") + "' WHERE BatteryModel='" + current["BatteryModel"].ToString() + "'";
                    cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }

                    // Also update the pci dataTable!
                    for (int i = 0; i < 16; i++)
                    {
                        if (((Main_Form)this.Owner).pci.Rows[i][0].ToString() == current["BatteryModel"].ToString())
                        {
                            // do we have multiple work orders?


                            string tempWOS = ((Main_Form)this.Owner).d.Rows[i][1].ToString();
                            char[] delims = { ' ' };
                            string[] A = tempWOS.Split(delims);

                            ((Main_Form)this.Owner).pci.Rows[i][0] = textBox2.Text.Replace("'", "''");
                            if (comboBox13.Text.Replace("'", "''") != "") { ((Main_Form)this.Owner).pci.Rows[i][1] = comboBox13.Text.Replace("'", "''"); }
                            else { ((Main_Form)this.Owner).pci.Rows[i][1] = "NiCd"; }
                            if (numericUpDown89.Text.Replace("'", "''") != "") { ((Main_Form)this.Owner).pci.Rows[i][2] = (float) GetDouble(numericUpDown89.Text.Replace("'", "''")); }      // negative 1 is the default...
                            else { ((Main_Form)this.Owner).pci.Rows[i][2] = "24"; }
                            if (numericUpDown88.Text.Replace("'", "''") != "") { ((Main_Form)this.Owner).pci.Rows[i][3] = (A.Length - 1) * (int) GetDouble(numericUpDown88.Text.Replace("'", "''")); }        // negative 1 is the default...
                            else { ((Main_Form)this.Owner).pci.Rows[i][3] = "-1"; }
                            if (textBox21.Text.Replace("'", "''") != "") { ((Main_Form)this.Owner).pci.Rows[i][4] = (float) GetDouble(textBox21.Text.Replace("'", "''")); }        // negative 1 is the default...
                            else { ((Main_Form)this.Owner).pci.Rows[i][4] = "-1"; }
                            if (textBox22.Text.Replace("'", "''") != "") { ((Main_Form)this.Owner).pci.Rows[i][5] = (float) GetDouble(textBox22.Text.Replace("'", "''")); }        // 24 is the default...
                            else { ((Main_Form)this.Owner).pci.Rows[i][5] = "-1"; }
                            if (textBox27.Text.Replace("'", "''") != "") { ((Main_Form)this.Owner).pci.Rows[i][6] = (float) GetDouble(textBox27.Text.Replace("'", "''")); }        // negative 1 is the default...
                            else { ((Main_Form)this.Owner).pci.Rows[i][6] = "-1"; }
                            if (textBox28.Text.Replace("'", "''") != "") { ((Main_Form)this.Owner).pci.Rows[i][7] = (float) GetDouble(textBox28.Text.Replace("'", "''")); }        // 1.75 is the default...
                            else { ((Main_Form)this.Owner).pci.Rows[i][7] = "1.75"; }
                            if (textBox29.Text.Replace("'", "''") != "") { ((Main_Form)this.Owner).pci.Rows[i][8] = (float) GetDouble(textBox29.Text.Replace("'", "''")); }        // negative 1 is the default...
                            else { ((Main_Form)this.Owner).pci.Rows[i][8] = "-1"; }
                        }
                    }

                    //now force an update on the binding by moving one ahead and then back...
                    toolStripCBBats.ComboBox.Text = textBox2.Text.Replace("'", "''");
                    MessageBox.Show(this, "Battery model " + textBox2.Text + "'s entry has been updated.", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    // we need to insert a new record...
                    // first check to see if the serial number is already in use.
                    string checkString = "SELECT * FROM BatteriesCustom WHERE BatteryModel = '" + textBox2.Text.Replace("'", "''") + "'";
                    DataSet checkSet = new DataSet();
                    OleDbCommand checkCmd = new OleDbCommand(checkString, conn);
                    OleDbDataAdapter checkAdapter = new OleDbDataAdapter(checkCmd);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        checkAdapter.Fill(checkSet);
                        conn.Close();
                    }

                    if (checkSet.Tables[0].Rows.Count > 0)
                    {
                        //we already have that serial number in the DB
                        // tell the user about that and return...
                        MessageBox.Show(this, "That battery model is already in the database!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        updateCurVals();
                        return;
                    }

                    // find the max value in the RecordID column so we know what to assign to the new record
                    int max;
                    try
                    {
                        max = CustomBats.Tables[0].AsEnumerable().Max(r => r.Field<int>("RecordID"));
                    }
                    catch
                    {
                        max = 0;
                    }
                    string cmdStr = "INSERT INTO BatteriesCustom (RecordID, BMFR, BatteryModel, BPN, BTECH, VOLT, NCELLS, CAP, " +
                       "BCVMIN, BCVMAX, COT, CCVMMIN, CCVMAX, CCAPV, NOTES, " +
                       "[T1Mode], T1Time1Hr, T1Time1Min, T1Curr1, T1Volts1, T1Time2Hr, T1Time2Min, T1Curr2, T1Volts2, " +
                       "[T2Mode], T2Time1Hr, T2Time1Min, T2Curr1, T2Volts1, T2Time2Hr, T2Time2Min, T2Curr2, T2Volts2, " +
                       "[T3Mode], T3Time1Hr, T3Time1Min, T3Curr1, T3Volts1, T3Time2Hr, T3Time2Min, T3Curr2, T3Volts2, " +
                       "[T4Mode], T4Time1Hr, T4Time1Min, T4Curr1, T4Volts1, T4Time2Hr, T4Time2Min, T4Curr2, T4Volts2, " +
                       "[T5Mode], T5Time1Hr, T5Time1Min, T5Curr1, T5Volts1, T5Time2Hr, T5Time2Min, T5Curr2, T5Volts2, " +
                       "[T6Mode], T6Time1Hr, T6Time1Min, T6Curr1, T6Volts1, T6Ohms, " +
                       "[T7Mode], T7Time1Hr, T7Time1Min, T7Curr1, T7Volts1, T7Ohms, " +
                       "[T8Mode], T8Time1Hr, T8Time1Min, T8Curr1, T8Volts1, T8Time2Hr, T8Time2Min, T8Curr2, T8Volts2, " +
                       "[T9Mode], T9Time1Hr, T9Time1Min, T9Curr1, T9Volts1, T9Time2Hr, T9Time2Min, T9Curr2, T9Volts2, " +
                       "[T10Mode], T10Time1Hr, T10Time1Min, T10Curr1, T10Volts1, T10Time2Hr, T10Time2Min, T10Curr2, T10Volts2, " +
                       "[T11Mode], T11Time1Hr, T11Time1Min, T11Curr1, T11Volts1, T11Ohms, " +
                       "[T12Mode], T12Time1Hr, T12Time1Min, T12Curr1, T12Volts1, T12Time2Hr, T12Time2Min, T12Curr2, T12Volts2, " +
                       "[T13Mode], T13Time1Hr, T13Time1Min, T13Curr1, T13Volts1, T13Time2Hr, T13Time2Min, T13Curr2, T13Volts2, " +
                       "[T14Mode], T14Time1Hr, T14Time1Min, T14Curr1, T14Volts1, T14Time2Hr, T14Time2Min, T14Curr2, T14Volts2, " +
                       "[T15Mode], T15Time1Hr, T15Time1Min, T15Curr1, T15Volts1, T15Time2Hr, T15Time2Min, T15Curr2, T15Volts2) " +
                        "VALUES (" + (max + 1).ToString() + ",'" +
                        textBox1.Text.Replace("'", "''") + "','" +
                        textBox2.Text.Replace("'", "''") + "','" +
                        textBox3.Text.Replace("'", "''") + "','" +
                        comboBox13.Text.Replace("'", "''") + "','" +
                        numericUpDown89.Text.Replace("'", "''") + "','" +
                        numericUpDown88.Text.Replace("'", "''") + "','" +
                        numericUpDown90.Text.Replace("'", "''") + "','" +
                        textBox21.Text.Replace("'", "''") + "','" +
                        textBox22.Text.Replace("'", "''") + "','" +
                        textBox23.Text.Replace("'", "''") + "','" +
                        textBox27.Text.Replace("'", "''") + "','" +
                        textBox28.Text.Replace("'", "''") + "','" +
                        textBox29.Text.Replace("'", "''") + "','" +
                        textBox36.Text.Replace("'", "''") + "','" +
                        // Full Charge-6 ("T1Mode, T1Time1Hr, T1Time1Min, T1Curr1, T1Volts1, T1Time2Hr, T1Time2Min, T1Curr2, T1Volts2, T1Ohms")
                        comboBox2.Text.Replace("'", "''") + "','" +
                        numericUpDown1.Text.Replace("'", "''") + "','" +
                        numericUpDown2.Text.Replace("'", "''") + "','" +
                        numericUpDown3.Text.Replace("'", "''") + "','" +
                        numericUpDown4.Text.Replace("'", "''") + "','" +
                        numericUpDown8.Text.Replace("'", "''") + "','" +
                        numericUpDown7.Text.Replace("'", "''") + "','" +
                        numericUpDown6.Text.Replace("'", "''") + "','" +
                        numericUpDown5.Text.Replace("'", "''") + "','" +
                        // Full Charge-4 ("T2Mode, T2Time1Hr, T2Time1Min, T2Curr1, T2Volts1, T2Time2Hr, T2Time2Min, T2Curr2, T2Volts2, T2Ohms")
                        comboBox1.Text.Replace("'", "''") + "','" +
                        numericUpDown16.Text.Replace("'", "''") + "','" +
                        numericUpDown15.Text.Replace("'", "''") + "','" +
                        numericUpDown14.Text.Replace("'", "''") + "','" +
                        numericUpDown13.Text.Replace("'", "''") + "','" +
                        numericUpDown12.Text.Replace("'", "''") + "','" +
                        numericUpDown11.Text.Replace("'", "''") + "','" +
                        numericUpDown10.Text.Replace("'", "''") + "','" +
                        numericUpDown9.Text.Replace("'", "''")  + "','" +
                        // Top Charge-4 ("T3Mode, T3Time1Hr, T3Time1Min, T1Curr1, T1Volts1, T1Time2Hr, T1Time2Min, T1Curr2, T1Volts2, T1Ohms")
                        comboBox3.Text.Replace("'", "''") + "','" +
                        numericUpDown24.Text.Replace("'", "''") + "','" +
                        numericUpDown23.Text.Replace("'", "''") + "','" +
                        numericUpDown22.Text.Replace("'", "''") + "','" +
                        numericUpDown21.Text.Replace("'", "''") + "','" +
                        numericUpDown20.Text.Replace("'", "''") + "','" +
                        numericUpDown19.Text.Replace("'", "''") + "','" +
                        numericUpDown18.Text.Replace("'", "''") + "','" +
                        numericUpDown17.Text.Replace("'", "''") + "','" +
                        // Top Charge-2 ("T4Mode, T4Time1Hr, T4Time1Min, T4Curr1, T4Volts1, T4Time2Hr, T4Time2Min, T4Curr2, T4Volts2, T4Ohms")
                        comboBox4.Text.Replace("'", "''") + "','" +
                        numericUpDown32.Text.Replace("'", "''") + "','" +
                        numericUpDown31.Text.Replace("'", "''") + "','" +
                        numericUpDown30.Text.Replace("'", "''") + "','" +
                        numericUpDown29.Text.Replace("'", "''") + "','" +
                        numericUpDown28.Text.Replace("'", "''") + "','" +
                        numericUpDown27.Text.Replace("'", "''") + "','" +
                        numericUpDown26.Text.Replace("'", "''") + "','" +
                        numericUpDown25.Text.Replace("'", "''") + "','" +
                        // Top Charge-1 ("T5Mode, T5Time1Hr, T5Time1Min, T5Curr1, T5Volts1, T5Time2Hr, T5Time2Min, T5Curr2, T5Volts2, T5Ohms")
                        comboBox5.Text.Replace("'", "''") + "','" +
                        numericUpDown40.Text.Replace("'", "''") + "','" +
                        numericUpDown39.Text.Replace("'", "''") + "','" +
                        numericUpDown38.Text.Replace("'", "''") + "','" +
                        numericUpDown37.Text.Replace("'", "''") + "','" +
                        numericUpDown36.Text.Replace("'", "''") + "','" +
                        numericUpDown35.Text.Replace("'", "''") + "','" +
                        numericUpDown34.Text.Replace("'", "''") + "','" +
                        numericUpDown33.Text.Replace("'", "''") + "','" +
                        // Capacity-1 ("T6Mode, T6Time1Hr, T6Time1Min, T6Curr1, T6Volts1, T6Time2Hr, T6Time2Min, T6Curr2, T6Volts2, T6Ohms")
                        comboBox6.Text.Replace("'", "''") + "','" +
                        numericUpDown45.Text.Replace("'", "''") + "','" +
                        numericUpDown44.Text.Replace("'", "''") + "','" +
                        numericUpDown43.Text.Replace("'", "''") + "','" +
                        numericUpDown42.Text.Replace("'", "''") + "','" +
                        numericUpDown41.Text.Replace("'", "''") + "','" +
                        // Discharge ("T7Mode, T7Time1Hr, T7Time1Min, T7Curr1, T7Volts1, T7Time2Hr, T7Time2Min, T7Curr2, T7Volts2, T7Ohms")
                        comboBox7.Text.Replace("'", "''") + "','" +
                        numericUpDown50.Text.Replace("'", "''") + "','" +
                        numericUpDown49.Text.Replace("'", "''") + "','" +
                        numericUpDown48.Text.Replace("'", "''") + "','" +
                        numericUpDown47.Text.Replace("'", "''") + "','" +
                        numericUpDown46.Text.Replace("'", "''") + "','" +
                        // Slow Charge-14 ("T8Mode, T8Time1Hr, T8Time1Min, T8Curr1, T8Volts1, T8Time2Hr, T8Time2Min, T8Curr2, T8Volts2, T8Ohms")
                        comboBox9.Text.Replace("'", "''") + "','" +
                        numericUpDown63.Text.Replace("'", "''") + "','" +
                        numericUpDown62.Text.Replace("'", "''") + "','" +
                        numericUpDown61.Text.Replace("'", "''") + "','" +
                        numericUpDown60.Text.Replace("'", "''") + "','" +
                        numericUpDown59.Text.Replace("'", "''") + "','" +
                        numericUpDown58.Text.Replace("'", "''") + "','" +
                        numericUpDown57.Text.Replace("'", "''") + "','" +
                        numericUpDown56.Text.Replace("'", "''") + "','" +
                        // Slow Charge-16 ("T9Mode, T9Time1Hr, T9Time1Min, T9Curr1, T9Volts1, T9Time2Hr, T9Time2Min, T9Curr2, T9Volts2, T9Ohms")
                        comboBox10.Text.Replace("'", "''") + "','" +
                        numericUpDown71.Text.Replace("'", "''") + "','" +
                        numericUpDown70.Text.Replace("'", "''") + "','" +
                        numericUpDown69.Text.Replace("'", "''") + "','" +
                        numericUpDown68.Text.Replace("'", "''") + "','" +
                        numericUpDown67.Text.Replace("'", "''") + "','" +
                        numericUpDown66.Text.Replace("'", "''") + "','" +
                        numericUpDown65.Text.Replace("'", "''") + "','" +
                        numericUpDown64.Text.Replace("'", "''") + "','" +
                        // Custom Chg ("T10Mode, T10Time1Hr, T10Time1Min, T10Curr1, T10Volts1, T10Time2Hr, T10Time2Min, T10Curr2, T10Volts2, T10Ohms")
                        comboBox11.Text.Replace("'", "''") + "','" +
                        numericUpDown79.Text.Replace("'", "''") + "','" +
                        numericUpDown78.Text.Replace("'", "''") + "','" +
                        numericUpDown77.Text.Replace("'", "''") + "','" +
                        numericUpDown76.Text.Replace("'", "''") + "','" +
                        numericUpDown75.Text.Replace("'", "''") + "','" +
                        numericUpDown74.Text.Replace("'", "''") + "','" +
                        numericUpDown73.Text.Replace("'", "''") + "','" +
                        numericUpDown72.Text.Replace("'", "''") + "','" +
                        // Custom Cap ("T11Mode, T11Time1Hr, T11Time1Min, T11Curr1, T11Volts1, T11Time2Hr, T11Time2Min, T11Curr2, T11Volts2, T11Ohms")
                        comboBox8.Text.Replace("'", "''") + "','" +
                        numericUpDown55.Text.Replace("'", "''") + "','" +
                        numericUpDown54.Text.Replace("'", "''") + "','" +
                        numericUpDown53.Text.Replace("'", "''") + "','" +
                        numericUpDown52.Text.Replace("'", "''") + "','" +
                        numericUpDown51.Text.Replace("'", "''") + "','" +
                        // Constant Voltage ("T12Mode, T12Time1Hr, T12Time1Min, T12Curr1, T12Volts1, T12Time2Hr, T12Time2Min, T12Curr2, T12Volts2, T12Ohms")
                        comboBox12.Text.Replace("'", "''") + "','" +
                        numericUpDown87.Text.Replace("'", "''") + "','" +
                        numericUpDown86.Text.Replace("'", "''") + "','" +
                        numericUpDown85.Text.Replace("'", "''") + "','" +
                        numericUpDown84.Text.Replace("'", "''") + "','" +
                        numericUpDown83.Text.Replace("'", "''") + "','" +
                        numericUpDown82.Text.Replace("'", "''") + "','" +
                        numericUpDown81.Text.Replace("'", "''") + "','" +
                        numericUpDown80.Text.Replace("'", "''") + "','" +
                        // Full Charge 4.5 ("T13Mode, T13Time1Hr, T13Time1Min, T13Curr1, T13Volts1, T13Time2Hr, T13Time2Min, T13Curr2, T13Volts2, T13Ohms")
                        comboBox14.Text.Replace("'", "''") + "','" +
                        numericUpDown98.Text.Replace("'", "''") + "','" +
                        numericUpDown97.Text.Replace("'", "''") + "','" +
                        numericUpDown96.Text.Replace("'", "''") + "','" +
                        numericUpDown95.Text.Replace("'", "''") + "','" +
                        numericUpDown94.Text.Replace("'", "''") + "','" +
                        numericUpDown93.Text.Replace("'", "''") + "','" +
                        numericUpDown92.Text.Replace("'", "''") + "','" +
                        numericUpDown91.Text.Replace("'", "''") +
                        // Custom Chg 2("T14Mode, T14Time1Hr, T14Time1Min, T14Curr1, T14Volts1, T14Time2Hr, T14Time2Min, T14Curr2, T14Volts2, T14Ohms")
                        comboBox15.Text.Replace("'", "''") + "','" +
                        numericUpDown106.Text.Replace("'", "''") + "','" +
                        numericUpDown105.Text.Replace("'", "''") + "','" +
                        numericUpDown104.Text.Replace("'", "''") + "','" +
                        numericUpDown103.Text.Replace("'", "''") + "','" +
                        numericUpDown102.Text.Replace("'", "''") + "','" +
                        numericUpDown101.Text.Replace("'", "''") + "','" +
                        numericUpDown100.Text.Replace("'", "''") + "','" +
                        numericUpDown99.Text.Replace("'", "''") + "','" +
                        // Custom Chg 3("T15Mode, T150Time1Hr, T15Time1Min, T15Curr1, T15Volts1, T15Time2Hr, T15Time2Min, T15Curr2, T15Volts2, T15Ohms")
                        comboBox16.Text.Replace("'", "''") + "','" +
                        numericUpDown114.Text.Replace("'", "''") + "','" +
                        numericUpDown113.Text.Replace("'", "''") + "','" +
                        numericUpDown112.Text.Replace("'", "''") + "','" +
                        numericUpDown111.Text.Replace("'", "''") + "','" +
                        numericUpDown110.Text.Replace("'", "''") + "','" +
                        numericUpDown109.Text.Replace("'", "''") + "','" +
                        numericUpDown108.Text.Replace("'", "''") + "','" +
                        numericUpDown107.Text.Replace("'", "''") + "','" +
                        // finished with inputs!
                        "')";

                    OleDbCommand cmd = new OleDbCommand(cmdStr, conn);
                    lock (Main_Form.dataBaseLock)
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    MessageBox.Show(this, "Battery model " + textBox2.Text + "'s entry has been created.", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    // update the dataTable with the new record ID also..
                    current[0] = max + 1;

                    bindingNavigatorAddNewItem.Enabled = true;

                }

                // also update the current vals..
                updateCurVals();
            }// end try
            catch (Exception ex)
            {
                MessageBox.Show(this,ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            // temp will store time in minutes
            int temp = 0;
            // we need to make sure that all the times add up to 6 hrs!

            if (numericUpDown1.Value == 6)
            {
                // 6 is the max
                numericUpDown2.Value = 0;
            }

            // figure out the first charge value and then calculate the second charge value
            temp = 60 * (int) numericUpDown1.Value + (int) numericUpDown2.Value;
            temp = 360 - temp;

            numericUpDown8.Value = (decimal) (temp / 60);
            numericUpDown7.Value = (decimal) (temp % 60);

        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            // temp will store time in minutes
            int temp = 0;
            // we need to make sure that all the times add up to 6 hrs!

            if (numericUpDown1.Value == 6)
            {
                // 6 is the max
                numericUpDown2.Value = 0;
            }

            // figure out the first charge value and then calculate the second charge value
            temp = 60 * (int)numericUpDown1.Value + (int)numericUpDown2.Value;
            temp = 360 - temp;

            numericUpDown8.Value = (decimal) (temp / 60);
            numericUpDown7.Value = (decimal)(temp % 60);
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "20 Dual Rate")
            {
                label39.Text = "Main Over Voltage";
            }
            else
            {
                label39.Text = "Peak Transfer Voltage";
            }
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "20 Dual Rate")
            {
                label18.Text = "Main Over Voltage";
            }
            else
            {
                label18.Text = "Peak Transfer Voltage";
            }
        }

        private void numericUpDown16_ValueChanged(object sender, EventArgs e)
        {
            // temp will store time in minutes
            int temp = 0;
            // we need to make sure that all the times add up to 6 hrs!

            if (numericUpDown16.Value == 0)
            {
                // 6 is the max
                numericUpDown15.Value = 0;
            }

            // figure out the first charge value and then calculate the second charge value
            temp = 60 * (int)numericUpDown16.Value + (int)numericUpDown15.Value;
            temp = 240 - temp;

            numericUpDown12.Value = (decimal)(temp / 60);
            numericUpDown11.Value = (decimal)(temp % 60);
        }

        private void numericUpDown15_ValueChanged(object sender, EventArgs e)
        {
            // temp will store time in minutes
            int temp = 0;
            // we need to make sure that all the times add up to 6 hrs!

            if (numericUpDown16.Value == 0)
            {
                // 6 is the max
                numericUpDown15.Value = 0;
            }

            // figure out the first charge value and then calculate the second charge value
            temp = 60 * (int)numericUpDown16.Value + (int)numericUpDown15.Value;
            temp = 240 - temp;

            numericUpDown12.Value = (decimal)(temp / 60);
            numericUpDown11.Value = (decimal)(temp % 60);
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown24_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown24.Value = 4;
        }

        private void numericUpDown32_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown32.Value = 2;
        }

        private void numericUpDown40_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown40.Value = 1;
        }

        private void comboBox5_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox5.Text == "11 Single Rate with Peak Stop")
            {
                label80.Text = "Peak Stop Voltage";
            }
            else
            {
                label80.Text = "Charge Over Voltage";
            }
        }

        private void comboBox4_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox4.Text == "11 Single Rate with Peak Stop")
            {
                label67.Text = "Peak Stop Voltage";
            }
            else
            {
                label67.Text = "Charge Over Voltage";
            }
        }

        private void comboBox3_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text == "11 Single Rate with Peak Stop")
            {
                label54.Text = "Peak Stop Voltage";
            }
            else
            {
                label54.Text = "Charge Over Voltage";
            }
        }

        private void comboBox6_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox6.Text == "31 Capacity Test")
            {
                label91.Visible = true;
                label92.Visible = true;
                numericUpDown43.Visible = true;
                label87.Visible = false;
                label88.Visible = false;
                numericUpDown41.Visible = false;
            }
            else
            {
                label91.Visible = false;
                label92.Visible = false;
                numericUpDown43.Visible = false;
                label87.Visible = true;
                label88.Visible = true;
                numericUpDown41.Visible = true;
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox9_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox9.Text == "11 Single Rate with Peak Stop")
            {
                label120.Text = "Peak Stop Voltage";
            }
            else
            {
                label120.Text = "Charge Over Voltage";
            }
        }

        private void comboBox11_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox11.Text == "10 Single Rate")
            {
                groupBox24.Visible = false;
                groupBox25.Text = "";
                label148.Text = "Charge Current";
                label146.Text = "Charge Over Voltage";
                if (numericUpDown75.Value == 0)
                {
                    numericUpDown75.Value = 0;
                }
            }
            else if (comboBox11.Text == "11 Single Rate with Peak Stop")
            {
                groupBox24.Visible = false;
                groupBox25.Text = "";
                label148.Text = "Charge Current";
                label146.Text = "Peak Stop Voltage";
            }
            else if(comboBox11.Text == "12 Constant Voltage")
            {
                groupBox24.Visible = false;
                groupBox25.Text = "";
                label148.Text = "Initial Current";
                label146.Text = "Charge Voltage";
            }
            else if(comboBox11.Text == "20 Dual Rate")
            {
                groupBox24.Visible = true;
                groupBox25.Text = "Main Charge";
                label148.Text = "Main Charge Current";
                label146.Text = "Main Over Voltage";
            }
            else
            {
                groupBox24.Visible = true;
                groupBox25.Text = "Main Charge";
                label148.Text = "Main Charge Current";
                label146.Text = "Peak Transfer Voltage";
            }
        }

        private void comboBox8_SelectedValueChanged(object sender, EventArgs e)
        {

            if (comboBox8.Text == "30 Full Discharge")
            {
                //resistance
                label106.Visible = false;
                label105.Visible = false;
                numericUpDown51.Visible = false;

                //voltage
                label108.Visible = false;
                label107.Visible = false;
                numericUpDown52.Visible = false;

                //current
                label110.Visible = true;
                label109.Visible = true;
                numericUpDown53.Visible = true;
            }
            else if (comboBox8.Text == "31 Capacity Test")
            {
                //resistance
                label106.Visible = false;
                label105.Visible = false;
                numericUpDown51.Visible = false;

                //voltage
                label108.Visible = true;
                label107.Visible = true;
                numericUpDown52.Visible = true;

                //current
                label110.Visible = true;
                label109.Visible = true;
                numericUpDown53.Visible = true;
            }
            else // (comboBox11.Text == "32 Constant Resistance")
            {
                //resistance
                label106.Visible = true;
                label105.Visible = true;
                numericUpDown51.Visible = true;

                //voltage
                label108.Visible = true;
                label107.Visible = true;
                numericUpDown52.Visible = true;

                //current
                label110.Visible = false;
                label109.Visible = false;
                numericUpDown53.Visible = false;

            }

        }

        private void comboBox10_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox10.Text == "11 Single Rate with Peak Stop")
            {
                label133.Text = "Peak Stop Voltage";
            }
            else
            {
                label133.Text = "Charge Over Voltage";
            }
        }

        private void comboBox13_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox13.Text == "Sealed Lead Acid")
            {
                ((Control)tabControl1.TabPages[0]).Enabled = false;
                ((Control)tabControl1.TabPages[1]).Enabled = false;
                ((Control)tabControl1.TabPages[7]).Enabled = false;
                ((Control)tabControl1.TabPages[8]).Enabled = false;
                ((Control)tabControl1.TabPages[9]).Enabled = false;
                label6.Visible = false;
                groupBox7.Visible = false;
                numericUpDown88.Visible = false;
                label32.Visible = true;
                label33.Visible = true;
                textBox21.Visible = true;
                textBox22.Visible = true;
            }
            else
            {
                ((Control)tabControl1.TabPages[0]).Enabled = true;
                ((Control)tabControl1.TabPages[1]).Enabled = true;
                ((Control)tabControl1.TabPages[7]).Enabled = true;
                ((Control)tabControl1.TabPages[8]).Enabled = true;
                ((Control)tabControl1.TabPages[9]).Enabled = true;
                label6.Visible = true;
                groupBox7.Visible = true;
                numericUpDown88.Visible = true;
                label32.Visible = false;
                label33.Visible = false;
                textBox21.Visible = false;
                textBox22.Visible = false;
            }
        }

        private void numericUpDown8_ValueChanged(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            if (toolStripCBBats.Text == "")
            {
                bindingNavigatorAddNewItem.Enabled = false;
            }
        }

        private void bindingNavigatorMovePreviousItem_Click(object sender, EventArgs e)
        {
            
            //remove the new record if there is one..
            if (bindingNavigatorAddNewItem.Enabled == false && bindingNavigator1.BindingSource.Position < CustomBats.Tables[0].Rows.Count)
            {
                CustomBats.Tables[0].Rows[CustomBats.Tables[0].Rows.Count - 1].Delete();
                bindingNavigatorAddNewItem.Enabled = true;
            }

        }

        private void bindingNavigatorMoveFirstItem_Click(object sender, EventArgs e)
        {
            //remove the new record if there is one..
            if (bindingNavigatorAddNewItem.Enabled == false && bindingNavigator1.BindingSource.Position < CustomBats.Tables[0].Rows.Count)
            {
                CustomBats.Tables[0].Rows[CustomBats.Tables[0].Rows.Count - 1].Delete();
                bindingNavigatorAddNewItem.Enabled = true;
            }
        }

        private void bindingNavigatorPositionItem_LocationChanged(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorPositionItem_TextChanged(object sender, EventArgs e)
        {

        }

        private void bindingNavigator1_LocationChanged(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && bindingNavigator1.BindingSource.Position < CustomBats.Tables[0].Rows.Count)
                {
                    CustomBats.Tables[0].Rows[CustomBats.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void toolStripCBBats_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //remove the new record if there is one..
                if (bindingNavigatorAddNewItem.Enabled == false && bindingNavigator1.BindingSource.Position < CustomBats.Tables[0].Rows.Count)
                {
                    CustomBats.Tables[0].Rows[CustomBats.Tables[0].Rows.Count - 1].Delete();
                    bindingNavigatorAddNewItem.Enabled = true;
                }
            }
            catch
            {
                //do nothing
            }
        }

        private void frmVECustomBats_Shown(object sender, EventArgs e)
        {
            updateCurVals();

            bindingNavigatorAddNewItem.PerformClick();
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

        private bool ValidateIt()
        {
            // do we need to validate?

            if (textBox2.Text != curTemp1 ||
                textBox1.Text != curTemp2 ||
                textBox3.Text != curTemp3 ||
                comboBox13.Text != curTemp4 ||

                numericUpDown89.Text != curTemp5 ||
                numericUpDown88.Text != curTemp6 ||
                numericUpDown90.Text != curTemp7 ||

                textBox27.Text != curTemp10 ||
                textBox28.Text != curTemp11 ||
                textBox29.Text != curTemp12 ||

                textBox21.Text != curTemp13 ||
                textBox22.Text != curTemp14 ||
                textBox23.Text != curTemp15 ||

                textBox36.Text != curTemp16 ||

                comboBox2.Text != curTemp17 ||

                numericUpDown1.Value != curTemp18 ||
                numericUpDown2.Value != curTemp19 ||
                numericUpDown3.Value != curTemp20 ||
                numericUpDown4.Value != curTemp21 ||

                numericUpDown8.Value != curTemp22 ||
                numericUpDown7.Value != curTemp23 ||
                numericUpDown6.Value != curTemp24 ||
                numericUpDown5.Value != curTemp25 ||

                comboBox1.Text != curTemp26 ||

                numericUpDown16.Value != curTemp27 ||
                numericUpDown15.Value != curTemp28 ||
                numericUpDown14.Value != curTemp29 ||
                numericUpDown13.Value != curTemp30 ||

                numericUpDown12.Value != curTemp31 ||
                numericUpDown11.Value != curTemp32 ||
                numericUpDown10.Value != curTemp33 ||
                numericUpDown9.Value != curTemp34 ||

                comboBox3.Text != curTemp35 ||

                numericUpDown24.Value != curTemp36 ||
                numericUpDown23.Value != curTemp37 ||
                numericUpDown22.Value != curTemp38 ||
                numericUpDown21.Value != curTemp39 ||

                numericUpDown20.Value != curTemp40 ||
                numericUpDown19.Value != curTemp41 ||
                numericUpDown18.Value != curTemp42 ||
                numericUpDown17.Value != curTemp43 ||

                comboBox4.Text != curTemp44 ||

                numericUpDown32.Value != curTemp45 ||
                numericUpDown31.Value != curTemp46 ||
                numericUpDown30.Value != curTemp47 ||
                numericUpDown29.Value != curTemp48 ||

                numericUpDown28.Value != curTemp49 ||
                numericUpDown27.Value != curTemp50 ||
                numericUpDown26.Value != curTemp51 ||
                numericUpDown25.Value != curTemp52 ||

                comboBox5.Text != curTemp53 ||
                
                numericUpDown40.Value != curTemp54 ||
                numericUpDown39.Value != curTemp55 ||
                numericUpDown38.Value != curTemp56 ||
                numericUpDown37.Value != curTemp57 ||

                numericUpDown36.Value != curTemp58 ||
                numericUpDown35.Value != curTemp59 ||
                numericUpDown34.Value != curTemp60 ||
                numericUpDown33.Value != curTemp61 ||

                comboBox12.Text != curTemp62 ||

                numericUpDown87.Value != curTemp63 ||
                numericUpDown86.Value != curTemp64 ||
                numericUpDown85.Value != curTemp65 ||
                numericUpDown84.Value != curTemp66 ||

                numericUpDown83.Value != curTemp67 ||
                numericUpDown82.Value != curTemp68 ||
                numericUpDown81.Value != curTemp69 ||
                numericUpDown80.Value != curTemp70 ||

                comboBox6.Text != curTemp71 ||

                numericUpDown45.Value != curTemp72 ||
                numericUpDown44.Value != curTemp73 ||
                numericUpDown43.Value != curTemp74 ||
                numericUpDown42.Value != curTemp75 ||
                numericUpDown41.Value != curTemp76 ||

                comboBox7.Text != curTemp77 ||

                numericUpDown50.Value != curTemp78 ||
                numericUpDown49.Value != curTemp79 ||
                numericUpDown48.Value != curTemp80 ||
                numericUpDown47.Value != curTemp81 ||
                numericUpDown46.Value != curTemp82 ||

                comboBox9.Text != curTemp83 ||

                numericUpDown63.Value != curTemp84 ||
                numericUpDown62.Value != curTemp85 ||
                numericUpDown61.Value != curTemp86 ||
                numericUpDown60.Value != curTemp87 ||

                numericUpDown59.Value != curTemp88 ||
                numericUpDown58.Value != curTemp89 ||
                numericUpDown57.Value != curTemp90 ||
                numericUpDown56.Value != curTemp101 ||

                comboBox10.Text != curTemp91 ||

                numericUpDown71.Value != curTemp92 ||
                numericUpDown70.Value != curTemp93 ||
                numericUpDown69.Value != curTemp94 ||
                numericUpDown68.Value != curTemp95 ||

                numericUpDown67.Value != curTemp96 ||
                numericUpDown66.Value != curTemp97 ||
                numericUpDown65.Value != curTemp98 ||
                numericUpDown64.Value != curTemp99 ||

                comboBox11.Text != curTemp100 ||

                numericUpDown79.Value != curTemp102 ||
                numericUpDown78.Value != curTemp103 ||
                numericUpDown77.Value != curTemp104 ||
                numericUpDown76.Value != curTemp105 ||

                numericUpDown75.Value != curTemp106 ||
                numericUpDown74.Value != curTemp107 ||
                numericUpDown73.Value != curTemp108 ||
                numericUpDown72.Value != curTemp109 ||

                comboBox8.Text != curTemp110 ||

                numericUpDown55.Value != curTemp111 ||
                numericUpDown54.Value != curTemp112 ||
                numericUpDown53.Value != curTemp113 ||
                numericUpDown52.Value != curTemp114 ||
                numericUpDown51.Value != curTemp115 ||
                
                curTemp116 != comboBox14.Text ||

                curTemp117 != numericUpDown96.Value ||
                curTemp118 != numericUpDown95.Value ||
                curTemp119 != numericUpDown92.Value ||
                curTemp120 != numericUpDown91.Value  ||
                
                comboBox15.Text != curTemp121 ||

                numericUpDown106.Value != curTemp122 ||
                numericUpDown105.Value != curTemp123 ||
                numericUpDown104.Value != curTemp124 ||
                numericUpDown103.Value != curTemp125 ||

                numericUpDown102.Value != curTemp126 ||
                numericUpDown101.Value != curTemp127 ||
                numericUpDown100.Value != curTemp128 ||
                numericUpDown99.Value != curTemp129 ||
                
                comboBox16.Text != curTemp130 ||

                numericUpDown114.Value != curTemp131 ||
                numericUpDown113.Value != curTemp132 ||
                numericUpDown112.Value != curTemp133 ||
                numericUpDown111.Value != curTemp134 ||

                numericUpDown110.Value != curTemp135 ||
                numericUpDown109.Value != curTemp136 ||
                numericUpDown108.Value != curTemp137 ||
                numericUpDown107.Value != curTemp138)
            {

                // they don't match!
                // ask if the user is sure that they want to continue...
                DialogResult dialogResult = MessageBox.Show(this, "Looks like this record has been updated without being saved.  Are you sure you want to navigate away without saving?", "Click Yes to continue or No stay here.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.No)
                {
                    updateCurVals();
                    return true;
                }
                else
                {
                    //sync everything..
                    updateCurVals();

                }
            }
            return false;
        }

        private void bindingNavigatorMovePreviousItem_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void bindingNavigatorMoveFirstItem_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void bindingNavigatorPositionItem_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void bindingNavigatorMoveNextItem_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void bindingNavigatorMoveLastItem_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void bindingNavigatorAddNewItem_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void bindingNavigatorDeleteItem_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void saveToolStripButton_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void frmVECustomBats_FormClosing(object sender, FormClosingEventArgs e)
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
                InhibitCB = true;
                Inhibit = true;
            }

        }

        private void bindingNavigator1_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void bindingNavigator1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            Inhibit = false;
            bindingNavigator1.Focus();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox1_Leave(object sender, EventArgs e)
        {

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

        private void tabPage10_Click(object sender, EventArgs e)
        {

        }

        private void groupBox28_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox29_Enter(object sender, EventArgs e)
        {

        }

        private void label185_Click(object sender, EventArgs e)
        {

        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox15_SelectedValueChanged(object sender, EventArgs e)
        {

            if (comboBox15.Text == "10 Single Rate")
            {
                groupBox28.Visible = false;
                groupBox29.Text = "";
                label181.Text = "Charge Current";
                label179.Text = "Charge Over Voltage";
                if (numericUpDown75.Value == 0)
                {
                    numericUpDown75.Value = 0;
                }
            }
            else if (comboBox15.Text == "11 Single Rate with Peak Stop")
            {
                groupBox28.Visible = false;
                groupBox29.Text = "";
                label181.Text = "Charge Current";
                label179.Text = "Peak Stop Voltage";
            }
            else if(comboBox15.Text == "12 Constant Voltage")
            {
                groupBox28.Visible = false;
                groupBox29.Text = "";
                label181.Text = "Initial Current";
                label179.Text = "Charge Voltage";
            }
            else if(comboBox15.Text == "20 Dual Rate")
            {
                groupBox28.Visible = true;
                groupBox29.Text = "Main Charge";
                label181.Text = "Main Charge Current";
                label179.Text = "Main Over Voltage";
            }
            else
            {
                groupBox28.Visible = true;
                groupBox29.Text = "Main Charge";
                label181.Text = "Main Charge Current";
                label179.Text = "Peak Transfer Voltage";
            }
        }

        private void comboBox16_SelectedValueChanged(object sender, EventArgs e)
        {
        

            if (comboBox16.Text == "10 Single Rate")
            {
                groupBox30.Visible = false;
                groupBox31.Text = "";
                label194.Text = "Charge Current";
                label192.Text = "Charge Over Voltage";
                if (numericUpDown75.Value == 0)
                {
                    numericUpDown75.Value = 0;
                }
            }
            else if (comboBox16.Text == "11 Single Rate with Peak Stop")
            {
                groupBox30.Visible = false;
                groupBox31.Text = "";
                label194.Text = "Charge Current";
                label192.Text = "Peak Stop Voltage";
            }
            else if(comboBox16.Text == "12 Constant Voltage")
            {
                groupBox30.Visible = false;
                groupBox31.Text = "";
                label194.Text = "Initial Current";
                label192.Text = "Charge Voltage";
            }
            else if(comboBox16.Text == "20 Dual Rate")
            {
                groupBox30.Visible = true;
                groupBox31.Text = "Main Charge";
                label194.Text = "Main Charge Current";
                label192.Text = "Main Over Voltage";
            }
            else
            {
                groupBox30.Visible = true;
                groupBox31.Text = "Main Charge";
                label194.Text = "Main Charge Current";
                label192.Text = "Peak Transfer Voltage";
            }

        }
        
    }
}
