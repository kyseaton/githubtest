﻿using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Xml.Serialization;
using System.IO;
using System.Deployment;

namespace NewBTASProto
{
    public partial class Splash : Form
    {
        private PictureBox pictureBox1;
        Timer tmr;

        bool licGood = false;
        // here are the valid licence numbers
        static string[] lics = new string[50] {
            "1111111111",                       // jfm
            "2222222222",                       // jfm
            "3333333333",                       // jfm
            "5339428555",                       // Chase @ Gulf Stream
            "6875678858",                       // Greg  @ Emerson
            "8595903103",
            "2013021954",                       // AOG
            "5388165955",                       // Bandi @ Scandinavian Avionics Norway
            "8303121362",                       // Benson Wong / Raymond Fung @ Topcast Hong Kong
            "5735222138",                       // Air New Zealand / Greg Donaldson
            
            "2995964807",                       // Lufthansa Technik Malta
            "7121966571",                       // Global Aerospace
            "2625219448",                       // Yardimcilar
            "9185399390",                       // Yardimcilar
            "1555832326",                       // John Burns @ Gulf Stream
            "6800782161",                       // Eric @ Will Air
            "9238872250",                       // Satair
            "2885597260",                       // Jon Ravenhall @ Satair
            "6346677964",                       // Rotortech Services
            "2328164969",                       // Aviation Parts EXE
            
            "8289656377",                       // Topcast Aviation
            "5241029672",                       // Saft Demo
            "4187028983",                       // Cook Aviation
            "7712998423",                       // New Zealand Air Defense 1
            "3844940117",                       // New Zealand Air Defense 2
            "5568154526",                       // New Zealand Air Defense 3
            "5900397404",                       // RB171 LLC
            "1924146292",                       // Satair Miami
            "4646305542",                       // Piedmont
            "6690621576",                       // TAP

            "4411606388",                       // Professional Technology Repairs
            "8799674435",                       // Air Iceland
            "8082544274",                       // L Brands
            "9007356233",                       // Piedmont
            "3230706410",                       // Interjet
            "8610841593",                       // Austral
            "4555182535",                       // Centurion Aviation
            "1098738068",                       // Satair
            "6529065984",                       // 
            "1439108005",                       // 
            
            "8278416243",                       // 
            "4526565694",                       // 
            "9075232581",                       // 
            "3201866946",                       // 
            "4873699807",                       // 
            "3115904672",                       // 
            "4806744307",                       // 
            "1492059149",                       // 
            "1917371840",                       // 
            "5999493295"                        // 
        };
        // this is the lic dataTable (here for its writexml and readxml methods
        DataTable lic = new DataTable();


        public Splash()
        {
            //if (System.Diagnostics.Process.GetProcessesByName(System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly().Location)).Length > 0) System.Diagnostics.Process.GetCurrentProcess().Kill();
            InitializeComponent();
        }

        public void Load_Globals()
        {
            string strAccessConn;
            string strAccessSelect;        
            OleDbConnection myAccessConn = null;

            //first we need to setup the folder string
            GlobalVars.folderString = Properties.Settings.Default.folderString;

            // make sure the program data location is set
            if (GlobalVars.folderString == "")
            {
                // reset it to the application folder path
                GlobalVars.folderString = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            }

            // create the connection
            try
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {

                MessageBox.Show(this, "Error: Failed to set up the database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;

            }

            //  open the db and pull in the options table
            try
            {
                strAccessSelect = @"SELECT * FROM Options";   
                DataSet options = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                try
                {
                    myAccessConn.Open();
                }
                catch
                {

                    //The DB isn't there!
                    //make the DB folder
                    try
                    {
                        System.IO.Directory.CreateDirectory(GlobalVars.folderString + @"\BTAS16_DB");
                    }
                    catch
                    {
                        // already there!
                    }

                    //now copy the file over
                    try
                    {
                        System.IO.File.WriteAllBytes(GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB", Properties.Resources.BTS16NV_clean);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Error: Failed to set up the database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    try
                    {
                        myAccessConn.Open();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Error: Failed to set up the database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

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
                MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                //MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                //MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                //MessageBox.Show(this, "Error: Failed to retrieve the required data from the DataBase. \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            // this section will load the noteservice settings
            //load the form settings
            try
            {
                NoteSet settings;
                XmlSerializer xs = new XmlSerializer(typeof(NoteSet));
                using (FileStream fs = new FileStream(GlobalVars.folderString + @"\BTAS16_DB\noteSet.xml", FileMode.Open))
                {
                    // This will read the XML from the file and create the new instance
                    // of settings
                    settings = xs.Deserialize(fs) as NoteSet;
                }

                // If the customer data was successfully deserialized we can transfer
                // the data from the to global vars.
                if (settings != null)
                {
                    GlobalVars.server = settings.server;
                    GlobalVars.port = settings.port;
                    GlobalVars.user = settings.user;
                    GlobalVars.pass = settings.pass;

                    GlobalVars.recipients = settings.recipients;
                    
                    GlobalVars.highLev = settings.highLev;
                    GlobalVars.medLev = settings.medLev;
                    GlobalVars.allLev = settings.allLev;

                    GlobalVars.stat0 = settings.stat0;
                    GlobalVars.stat1 = settings.stat1;
                    GlobalVars.stat2 = settings.stat2;
                    GlobalVars.stat3 = settings.stat3;
                    GlobalVars.stat4 = settings.stat4;
                    GlobalVars.stat5 = settings.stat5;
                    GlobalVars.stat6 = settings.stat6;
                    GlobalVars.stat7 = settings.stat7;
                    GlobalVars.stat8 = settings.stat8;
                    GlobalVars.stat9 = settings.stat9;
                    GlobalVars.stat10 = settings.stat10;
                    GlobalVars.stat11 = settings.stat11;
                    GlobalVars.stat12 = settings.stat12;
                    GlobalVars.stat13 = settings.stat13;
                    GlobalVars.stat14 = settings.stat14;
                    GlobalVars.stat15 = settings.stat15;

                    GlobalVars.all = settings.all;

                    GlobalVars.noteOn = settings.on;

                }
            }// end try
            catch
            {
                // we aren't sending anything out, so turn off the notification service.
                GlobalVars.noteOn = false;
            }

            //final step is to make sure we have a workorder logo.  If we don't we need to copy the one from resources to the correct location...
            try
            {
                if (!File.Exists(GlobalVars.folderString + @"\BTAS16_DB\rp_logo.jpg"))
                {
                    Properties.Resources.rp_logo.Save(GlobalVars.folderString + @"\BTAS16_DB\rp_logo.jpg");
                }
            }
            catch
            {

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

            //load the licence xml file and compare it to the licence list

            try
            {
                //create the columns
                lic.Columns.Add("num", typeof(string));
                //name the table
                lic.TableName = "lic_file";
                //now read in what we got!
                lic.ReadXml(GlobalVars.folderString + @"\BTAS16_DB\lic_file.xml");

                //now try to compare the lic # to the approved list...
                foreach (string posLic in lics)
                {
                    if (posLic == lic.Rows[0][0].ToString())
                    {
                        licGood = true;
                        break;
                    }
                }// end foreach
            }
            catch
            {
                //the licence file is not there!
                lic.Clear();
                licGood = false;
            }



        }



        void tmr_Tick(object sender, EventArgs e)
        {
            //after 3 sec stop the timer
            tmr.Stop();

            if (licGood == false)
            {
                //Ask for a licence number here!
                
                string temp = Microsoft.VisualBasic.Interaction.InputBox("The licence key wasn't found.  Please enter your licence key number.", "Licence not found!");

                //check the number
                //now try to compare the lic # to the approved list...
                foreach (string posLic in lics)
                {
                    if (posLic == temp)
                    {
                        licGood = true;
                        break;
                    }
                }// end foreach

                // if the number is bad, we'll return
                if(licGood == false)
                {
                    MessageBox.Show(this, "Cannot continue without a good licence key.  Contact JFM if you need a valid licence key.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Dispose();
                }
                //otherwise let's write the number to file and continue
                else
                {
                    lic.Rows.Add();
                    lic.Rows[0][0] = temp;
                    lic.WriteXml(GlobalVars.folderString + @"\BTAS16_DB\lic_file.xml");
                }

            }

            //display mainform

            try
            {
                Main_Form mf = new Main_Form();
                // update the options menu
                mf.Owner = this;
                mf.Show();
                //hide this form
                this.Hide();
            }
            catch(Exception ex)
            {
                MessageBox.Show("From Main Form: " + ex.Message + Environment.NewLine + ex.StackTrace);
            }

        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Splash));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::NewBTASProto.Properties.Resources.splash6_K;
            this.pictureBox1.Location = new System.Drawing.Point(-3, -1);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(656, 352);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // Splash
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(653, 351);
            this.Controls.Add(this.pictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(8, 9, 8, 9);
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
        public void checkDB()
        {

            // USE this function to update the DB!

            //////////////////////////////////////// Look to see if we have the AVE column in the WaterLevel table//////////

            // check to see if we have a "BatteryProfiles" table and create it if we don't
            string strAccessConn;
            string strAccessSelect;
            OleDbConnection myAccessConn;

            // create the connection
            try
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalVars.folderString + @"\BTAS16_DB\BTS16NV.MDB";
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //  open the db and try to get Ave out of the "WaterLevel" table
            try
            {
                strAccessSelect = @"SELECT AVE FROM WaterLevel";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

            }
            catch (Exception ex)
            {

                if (ex is OleDbException)
                {
                    // we didn't find the Ave field, so we need to create it!
                    // we need to  alter table
                    string strUpdateCMD = "ALTER TABLE WaterLevel ADD AVE Number";
                    OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();

                }
                else {
                    myAccessConn.Close();
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            //Now lets see if we have the combination test table and add it if we don't
            try
            {
                strAccessSelect = @"SELECT Steps FROM ComboTest";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

            }
            catch (Exception ex)
            {
                myAccessConn.Close();

                if (ex is OleDbException)
                {
                    // we didn't find the ComboTest Table
                    // lets add it!

                    try
                    {
                        strAccessSelect = "CREATE TABLE ComboTest (CTID AUTOINCREMENT PRIMARY KEY, ComboTestName Text, Steps Number, Step1 Text, Step2 Text, Step3 Text, Step4 Text, Step5 Text, Step6 Text, Step7 Text, Step8 Text, Step9 Text, Step10 Text, Step11 Text, Step12 Text, Step13 Text, Step14 Text, Step15 Text)";
                        OleDbCommand cmd = new OleDbCommand(strAccessSelect, myAccessConn);
                        myAccessConn.Open();
                        cmd.ExecuteNonQuery();
                        myAccessConn.Close();

                        strAccessSelect = "INSERT INTO ComboTest (ComboTestName, Steps, Step1, Step2) VALUES ('FC-6 Cap-1', 2, 'Full Charge-6', 'Capacity-1')";
                        cmd = new OleDbCommand(strAccessSelect, myAccessConn);
                        myAccessConn.Open();
                        cmd.ExecuteNonQuery();
                        myAccessConn.Close();

                        strAccessSelect = "INSERT INTO ComboTest (ComboTestName, Steps, Step1, Step2) VALUES ('FC-4 Cap-1', 2, 'Full Charge-4', 'Capacity-1')";
                        cmd = new OleDbCommand(strAccessSelect, myAccessConn);
                        myAccessConn.Open();
                        cmd.ExecuteNonQuery();
                        myAccessConn.Close();
                    }
                    catch { myAccessConn.Close(); }
                }
                else
                {
                    myAccessConn.Close();
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // we also need to make sure we have the wait and temp columns
            try
            {
                strAccessSelect = @"SELECT Wait AND Time FROM ComboTest";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

            }
            catch (Exception ex)
            {
                myAccessConn.Close();

                if (ex is OleDbException)
                {
                    // we didn't find the Wait and Time columns in the ComboTest Table
                    // lets add them!

                    try
                    {
                        strAccessSelect = "ALTER TABLE ComboTest ADD WaitTime Number, TempMar Number";
                        OleDbCommand cmd = new OleDbCommand(strAccessSelect, myAccessConn);
                        myAccessConn.Open();
                        cmd.ExecuteNonQuery();
                        myAccessConn.Close();
                    }
                    catch { myAccessConn.Close(); }
                }
                else
                {
                    myAccessConn.Close();
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
 

            //  open the db and check to see if we have the 4.5 hours settings in the DB
            try
            {
                strAccessSelect = @"SELECT T13Mode FROM BatteriesCustom";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

            }
            catch (Exception ex)
            {

                if (ex is OleDbException)
                {
                    // we didn't find the table, so we need to create it!
                    // we need to  insert
                    string strUpdateCMD = "ALTER TABLE BatteriesCustom ADD T13Mode Text(255), T13Time1Hr Text(255), T13Time1Min Text(255), T13Curr1 Text(255), T13Volts1 Text(255), T13Time2Hr Text(255), T13Time2Min Text(255), T13Curr2 Text(255), T13Volts2 Text(255), T13Ohms Text(255)";
                    OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();

                }
                else
                {
                    myAccessConn.Close();
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            //  open the db and check to see if we have the custom charge #2 and #3 in the DB
            try
            {
                strAccessSelect = @"SELECT T14Mode FROM BatteriesCustom";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

            }
            catch (Exception ex)
            {

                if (ex is OleDbException)
                {
                    // we didn't find the table, so we need to create it!
                    // we need to  insert
                    string strUpdateCMD = "ALTER TABLE BatteriesCustom ADD T14Mode Text(255), T14Time1Hr Text(255), T14Time1Min Text(255), T14Curr1 Text(255), T14Volts1 Text(255), T14Time2Hr Text(255), T14Time2Min Text(255), T14Curr2 Text(255), T14Volts2 Text(255), T14Ohms Text(255), T15Mode Text(255), T15Time1Hr Text(255), T15Time1Min Text(255), T15Curr1 Text(255), T15Volts1 Text(255), T15Time2Hr Text(255), T15Time2Min Text(255), T15Curr2 Text(255), T15Volts2 Text(255), T15Ohms Text(255)";
                    OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();

                }
                else
                {
                    myAccessConn.Close();
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            //  open the db and check to see if we have the As Recieved test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'As Received'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find full charge 4.5 in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('As Received', 3, 2)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the Full Charge-6 test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Full Charge-6'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find full charge 4.5 in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Full Charge-6', 73, 300)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the Full Charge-4 test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Full Charge-4'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find full charge 4.5 in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Full Charge-4', 61, 240)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the 4.5 hours test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Full Charge-4.5'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find full charge 4.5 in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Full Charge-4.5', 55, 300)" ;
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }


            catch (Exception ex)
            {
                
                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
            }

            //  open the db and check to see if we have the Top Charge-4 test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Top Charge-4'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find top charge 4 in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Top Charge-4', 61, 240)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the Top Charge-2 test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Top Charge-2'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find top charge 2 in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Top Charge-2', 41, 180)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the Top Charge-1 test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Top Charge-1'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find top charge 1 in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Top Charge-1', 61, 60)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the Constant Voltage test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Constant Voltage'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find constant voltage in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Constant Voltage', 73, 300)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the Capacity-1 test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Capacity-1'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find capactiy 1 in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Capacity-1', 61, 60)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the Discharge test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Discharge'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find Discharge in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Discharge', 61, 60)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the Slow Charge-16 test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Slow Charge-16'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find slow charge 16 in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Slow Charge-16', 61, 960)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the Slow Charge-14 test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Slow Charge-14'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find slow charge 14 in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Slow Charge-14', 73, 720)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the Custom Chg test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Custom Chg'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find custom charge in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Custom Chg', 60, 60)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the Custom Chg 2 test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Custom Chg 2'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find custom charge 2 in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Custom Chg 2', 60, 60)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the Custom Chg 3 test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Custom Chg 3'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find custom charge 3 in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Custom Chg 3', 60, 60)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }


            //  open the db and check to see if we have the Custom Cap test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Custom Cap'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find custom cap in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Custom Cap', 60, 60)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the Shorting-16 test settings in the DB
            try
            {
                strAccessSelect = @"SELECT * FROM TestType WHERE TESTNAME = 'Shorting-16'";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

                // access for 0 to test if anything is there...
                if (test.Tables[0].Rows.Count < 1)
                {
                    // we didn't find Shorting-16 in the table, so we need to add it!
                    // we need to  insert
                    string strUpdateCMD = "INSERT INTO TestType ([TESTNAME], Readings, [Interval]) VALUES ('Shorting-16', 61, 960)";
                    myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessConn.Open();
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();
                }

            }
            catch (Exception ex)
            {

                myAccessConn.Close();
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            //  open the db and check to see if we have the test setting fields in the testtype table
            try
            {
                strAccessSelect = @"SELECT TMode FROM TestType";
                DataSet test = new DataSet();
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(test, "test");
                myAccessConn.Close();

            }
            catch (Exception ex)
            {

                if (ex is OleDbException)
                {
                    // we didn't find the table, so we need to create it!
                    // we need to  insert
                    string strUpdateCMD = "ALTER TABLE TestType ADD TMode Text(255), TTime1Hr Text(255), TTime1Min Text(255), TCurr1 Text(255), TVolts1 Text(255), TTime2Hr Text(255), TTime2Min Text(255), TCurr2 Text(255), TVolts2 Text(255), TOhms Text(255), TTimeDHr Text(255), TTimeDMin Text(255), TCurrD Text(255), TVoltsD Text(255)";
                    OleDbCommand myAccessCommand = new OleDbCommand(strUpdateCMD, myAccessConn);
                    myAccessCommand.ExecuteNonQuery();
                    myAccessConn.Close();

                }
                else
                {
                    myAccessConn.Close();
                    MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            
        }




    }


}
