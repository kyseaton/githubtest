using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;

namespace NewBTASProto
{
    public partial class Splash : Form
    {
        private PictureBox pictureBox1;
        Timer tmr;

        public Splash()
        {
            InitializeComponent();
        }

        private void Load_Globals()
        {
            string strAccessConn;
            string strAccessSelect;        
            OleDbConnection myAccessConn;

            // create the connection
            try
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
                return;
            }

            //  open the db and pull in the options table
            try
            {
                strAccessSelect = @"SELECT * FROM Options";   
                DataSet options = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(options, "Options");

                // use the information to set the globals
                if (options.Tables[0].Rows[0][0].ToString() == "F.") { GlobalVars.useF = true; }
                else { GlobalVars.useF = false; }

                if (options.Tables[0].Rows[0][1].ToString() == "Pos. to Neg.") { GlobalVars.Pos2Neg = true; }
                else { GlobalVars.Pos2Neg = false; }

                GlobalVars.businessName = options.Tables[0].Rows[0][2].ToString();

                GlobalVars.highlightCurrent = false;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                myAccessConn.Close();
                return;
            }

            //  now for the comport settings
            try
            {
                // Load the Comconfig table...
                strAccessSelect = @"SELECT * FROM Comconfig";
                DataSet comconfig = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myDataAdapter.Fill(comconfig, "Comconfig");
                GlobalVars.CSCANComPort = comconfig.Tables[0].Rows[0][0].ToString();
                GlobalVars.ICComPort = comconfig.Tables[0].Rows[0][1].ToString();

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                myAccessConn.Close();
                return;
            }


            // Now We'll pull in the  AutoConfig setting which is located in ProgramSettings
            try
            {
                // Load the Comconfig table...
                strAccessSelect = "SELECT * FROM ProgramSettings WHERE SettingName='AutoConfigCharger'";
                DataSet progSettings = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myDataAdapter.Fill(progSettings, "progSettings");
                if (progSettings.Tables[0].Rows.Count < 1)
                {
                    // we need to insert
                    string strUpdateCMD = "INSERT INTO ProgramSettings (SettingName,SettingValue) VALUES ('AutoConfigCharger','false');";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessCommand.ExecuteNonQuery();
                    // don't forget to set the global!
                    GlobalVars.autoConfig = false;
                }
                else
                {
                    // read what we got!
                    if ((string) progSettings.Tables[0].Rows[0][2] == "False")
                    {
                        GlobalVars.autoConfig = false;
                    }
                    else
                    {
                        GlobalVars.autoConfig = true;
                    }

                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                myAccessConn.Close();
                return;
            }

            // Now We'll pull in the current tech setting which is located in ProgramSettings
            try
            {
                // Load the Comconfig table...
                strAccessSelect = "SELECT * FROM ProgramSettings WHERE SettingName='CurrentTech'";
                DataSet progSettings = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myDataAdapter.Fill(progSettings, "progSettings");
                if (progSettings.Tables[0].Rows.Count < 1)
                {
                    // we need to insert
                    string strUpdateCMD = "INSERT INTO ProgramSettings (SettingName,SettingValue) VALUES ('CurrentTech','Technician');";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessCommand.ExecuteNonQuery();
                    // don't forget to set the global!
                    GlobalVars.currentTech = "Technician";
                }
                else
                {
                    // read what we got!
                    GlobalVars.currentTech = (string) progSettings.Tables[0].Rows[0][2];
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error: Failed to retrieve the required data from the DataBase.\n" + ex.Message);
                myAccessConn.Close();
                return;
            }


            finally
            {
                myAccessConn.Close();
            }

            //This section is here to initialize the ICSettings array, which will later be filled either via the manual interface or the automatic charger programming routine
            for (int num = 0; num < 16; num++)
            {
                GlobalVars.ICSettings[num] = new ICSettingStore(num);
            }

            
        }


        private void SplashScreen_Shown(object sender, EventArgs e)
        {
            Load_Globals();
            checkDB();
            tmr = new Timer();
            //set time interval 3 sec
            tmr.Interval = 3000;
            //starts the timer
            tmr.Start();
            tmr.Tick += tmr_Tick;
        }



        void tmr_Tick(object sender, EventArgs e)
        {
            //after 3 sec stop the timer
            tmr.Stop();
            //display mainform
            Main_Form mf = new Main_Form();
            // update the options menu
            mf.Show();
            //hide this form
            this.Hide();
        }

        private void InitializeComponent()
        {
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::NewBTASProto.Properties.Resources.splash6;
            this.pictureBox1.Location = new System.Drawing.Point(-2, -1);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(491, 286);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // Splash
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(490, 285);
            this.Controls.Add(this.pictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(6, 7, 6, 7);
            this.Name = "Splash";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Splash Screen";
            this.Shown += new System.EventHandler(this.SplashScreen_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        // this function will check the DB and update it accordingly
        private void checkDB()
        {

            // USE this function to update the DB!


            //// check to see if we have a "BatteryProfiles" table and create it if we don't
            //string strAccessConn;
            //string strAccessSelect;
            //OleDbConnection myAccessConn;

            //// create the connection
            //try
            //{
            //    strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Kyle\Documents\Visual Studio 2013\Projects\NewBTASProto\BTS16NV.MDB";
            //    myAccessConn = new OleDbConnection(strAccessConn);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Error: Failed to create a database connection. \n" + ex.Message);
            //    return;
            //}

            ////  open the db and try to get something out of the "BatteryProfiles" table
            //try
            //{
            //    strAccessSelect = @"SELECT FIRST FROM BatteryProfiles;";
            //    DataSet test = new DataSet();
            //    OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
            //    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

            //    myAccessConn.Open();
            //    myDataAdapter.Fill(test, "test");

            //}
            //catch (Exception ex)
            //{
            //    myAccessConn.Close();
            //    MessageBox.Show("You do not have a battery Profiles Table. Please convert to the new database!");
            //    //if (ex is OleDbException)
            //    //{
            //    //    // we didn't find the table, so we need to create it!
            //    //    // we need to insert
            //    //    string strUpdateCMD = "CREATE TABLE BatteryProfiles (RecordID AutoNumber,Model Text,NCELLS Integer,MaxCellVoltage Single," +
            //    //        "BMFR Text,BPN Text,BTECH Text,Notes Memo," +
            //    //    ");";
            //    //    OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
            //    //    myAccessCommand.ExecuteNonQuery();

            //    //}
            //    //else {throw ex;}
            //}
            
        }


    }


}
