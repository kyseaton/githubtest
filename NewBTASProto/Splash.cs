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
                GlobalVars.CSCANComPort = "COM" + comconfig.Tables[0].Rows[0][0].ToString();
                GlobalVars.ICComPort = "COM" + comconfig.Tables[0].Rows[0][1].ToString();

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
            
        }


        private void SplashScreen_Shown(object sender, EventArgs e)
        {
            Load_Globals();
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

    }


}
