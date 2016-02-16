using System;
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
        static string[] lics = new string[30] {
            "1111111111",                       // jfm
            "2222222222",                       // jfm
            "3333333333",                       // jfm
            "5339428555",                       // Chase @ Gulf Stream
            "6875678858",                       // Greg  @ Emerson
            "8595903103",
            "2013021954",
            "5388165955",
            "8303121362",
            "5735222138",
            
            "2995964807",
            "7121966571",
            "2625219448",
            "9185399390",
            "1555832326",
            "6800782161",
            "9238872250",
            "2885597260",
            "6346677964",
            "2328164969",
            
            "8289656377",
            "5241029672",
            "4187028983",
            "7712998423",
            "3844940117",
            "5568154526",
            "5900397404",
            "1924146292",
            "4646305542",
            "6690621576"
        };
        // this is the lic dataTable (here for its writexml and readxml methods
        DataTable lic = new DataTable();


        public Splash()
        {
            //if (System.Diagnostics.Process.GetProcessesByName(System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly().Location)).Length > 0) System.Diagnostics.Process.GetCurrentProcess().Kill();
            InitializeComponent();
        }

        private void Load_Globals()
        {
            string strAccessConn;
            string strAccessSelect;        
            OleDbConnection myAccessConn = null;

            // create the connection
            try
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
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
                        System.IO.Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB");
                    }
                    catch
                    {
                        // already there!
                    }

                    //now copy the file over
                    try
                    {
                        System.IO.File.WriteAllBytes(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB", Properties.Resources.BTS16NV_clean);
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
                using (FileStream fs = new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\noteSet.xml", FileMode.Open))
                {
                    // This will read the XML from the file and create the new instance
                    // of CustomerData
                    settings = xs.Deserialize(fs) as NoteSet;
                }

                // If the customer data was successfully deserialized we can transfer
                // the data from the instance to the form.
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
                    GlobalVars.noteOff = settings.off;

                }
            }// end try
            catch
            {
                // do nothing...
            }

            //final step is to make sure we have a workorder logo.  If we don't we need to copy the one from resources to the correct location...
            try
            {
                if (!File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\rp_logo.jpg"))
                {
                    Properties.Resources.rp_logo.Save(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\rp_logo.jpg");
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

            //load the licence xml file and comare it to the licence list

            try
            {
                //create the columns
                lic.Columns.Add("num", typeof(string));
                //name the table
                lic.TableName = "lic_file";
                //now read in what we got!
                lic.ReadXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\lic_file.xml");

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
                    MessageBox.Show(this, "Cannot continue without a good licence key.  Contact JFM if you need a valid licence key.");
                    this.Dispose();
                }
                //otherwise let's write the number to file and continue
                else
                {
                    lic.Rows.Add();
                    lic.Rows[0][0] = temp;
                    lic.WriteXml(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\lic_file.xml");
                }

            }

            //display mainform

            try
            {
                Main_Form mf = new Main_Form();
                // update the options menu
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
            this.pictureBox1.Location = new System.Drawing.Point(-2, -1);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(492, 286);
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
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
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

            //////////////////////////////////////// Look to see if we have the AVE column in the WaterLevel table//////////

            // check to see if we have a "BatteryProfiles" table and create it if we don't
            string strAccessConn;
            string strAccessSelect;
            OleDbConnection myAccessConn;

            // create the connection
            try
            {
                strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\BTS16NV.MDB";
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error: Failed to create a database connection.  \n" + ex.Message + "\n" + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //  open the db and try to get something out of the "BatteryProfiles" table
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
                    // we didn't find the table, so we need to create it!
                    // we need to insert
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
            
        }


    }


}
