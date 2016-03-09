using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using System.Threading;
using System.Xml.Serialization;
using System.IO;

namespace NewBTASProto
{
    public partial class NoteServe : Form
    {
        public NoteServe()
        {
            InitializeComponent();

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
                    textBox1.Text = settings.server;
                    textBox2.Text = settings.port;
                    textBox3.Text = settings.user;
                    textBox4.Text = settings.pass;

                    char[] delims = { ',' };

                    foreach(string s in settings.recipients.Split(delims))
                    {
                        if (s != "")
                        {
                            listBox1.Items.Add(s);
                        }
                    }

                    radioButton1.Checked = settings.highLev;
                    radioButton2.Checked = settings.medLev;
                    radioButton3.Checked = settings.allLev;

                    checkBox1.Checked = settings.stat0;
                    checkBox2.Checked = settings.stat1;
                    checkBox3.Checked = settings.stat2;
                    checkBox4.Checked = settings.stat3;
                    checkBox5.Checked = settings.stat4;
                    checkBox6.Checked = settings.stat5;
                    checkBox7.Checked = settings.stat6;
                    checkBox8.Checked = settings.stat7;
                    checkBox9.Checked = settings.stat8;
                    checkBox10.Checked = settings.stat9;
                    checkBox11.Checked = settings.stat10;
                    checkBox12.Checked = settings.stat11;
                    checkBox13.Checked = settings.stat12;
                    checkBox14.Checked = settings.stat13;
                    checkBox15.Checked = settings.stat14;
                    checkBox16.Checked = settings.stat15;

                    checkBox17.Checked = settings.all;

                    radioButton4.Checked = settings.on;
                    radioButton5.Checked = settings.off;
                }
            }// end try
            catch
            {
                // do nothing...
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //save settings...
            // Create an instance of the CustomerData class and populate
            // it with the data from the form.
            NoteSet settings = new NoteSet();
            
            settings.server = textBox1.Text;
            settings.port = textBox2.Text;
            settings.user = textBox3.Text;            
            settings.pass = textBox4.Text;

            foreach (var row in listBox1.Items)
            {
                if (row.ToString() != "")
                {
                    settings.recipients += row.ToString() + ",";
                }
            }

            settings.highLev = radioButton1.Checked;
            settings.medLev = radioButton2.Checked;
            settings.allLev = radioButton3.Checked;
            
            settings.stat0 = checkBox1.Checked;
            settings.stat1 = checkBox2.Checked;
            settings.stat2 = checkBox3.Checked;
            settings.stat3 = checkBox4.Checked;
            settings.stat4 = checkBox5.Checked;
            settings.stat5 = checkBox6.Checked;
            settings.stat6 = checkBox7.Checked;
            settings.stat7 = checkBox8.Checked;
            settings.stat8 = checkBox9.Checked;
            settings.stat9 = checkBox10.Checked;
            settings.stat10 = checkBox11.Checked;
            settings.stat11 = checkBox12.Checked;
            settings.stat12 = checkBox13.Checked;
            settings.stat13 = checkBox14.Checked;
            settings.stat14 = checkBox15.Checked;
            settings.stat15 = checkBox16.Checked;

            settings.all = checkBox17.Checked;

            settings.on = radioButton4.Checked;
            settings.off = radioButton5.Checked;

            //Create and XmlSerializer to serialize the data to a file
            XmlSerializer xs = new XmlSerializer(typeof(NoteSet));
            using (FileStream fs = new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\BTAS16_DB\noteSet.xml", FileMode.Create))
            {
                xs.Serialize(fs, settings);
            }

            //now save this to the globals...
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

            this.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {

            // do everything on a helper thread...
            ThreadPool.QueueUserWorkItem(s =>
            {

                try
                {
                    // Create a System.Net.Mail.MailMessage object
                    MailMessage message = new MailMessage();

                    // Add a recipient
                    foreach (var row in listBox1.Items)
                    {
                        if (row.ToString() != "")
                        {
                            message.To.Add(row.ToString().Trim());
                        }
                    }
                    

                    // Add a message subject
                    message.Subject = "Test BTAS Message";

                    // Add a message body
                    message.Body = "This is a test BTAS message.";

                    // Create a System.Net.Mail.MailAddress object and 
                    // set the sender email address and display name.
                    message.From = new MailAddress(textBox3.Text);

                    // Create a System.Net.Mail.SmtpClient object
                    // and set the SMTP host and port number
                    SmtpClient smtp = new SmtpClient(textBox1.Text, int.Parse(textBox2.Text));

                    // If your server requires authentication add the below code
                    // =========================================================
                    // Enable Secure Socket Layer (SSL) for connection encryption
                    smtp.EnableSsl = true;

                    // Do not send the DefaultCredentials with requests
                    smtp.UseDefaultCredentials = false;

                    // Create a System.Net.NetworkCredential object and set
                    // the username and password required by your SMTP account
                    smtp.Credentials = new System.Net.NetworkCredential(textBox3.Text, textBox4.Text);
                    // =========================================================

                    // Send the message
                    smtp.Send(message);
                    MessageBox.Show(this, "Test Message Sent!", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

            });
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //add in the entry, but make sure something has been entered into the text box...
            if (textBox5.Text != "")
            {
                listBox1.Items.Add(textBox5.Text);
                textBox5.Text = "";
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Remove(listBox1.SelectedItem);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox17.Checked == true){
                checkBox1.Checked = true;
                checkBox2.Checked = true;
                checkBox3.Checked = true;
                checkBox4.Checked = true;
                checkBox5.Checked = true;
                checkBox6.Checked = true;
                checkBox7.Checked = true;
                checkBox8.Checked = true;
                checkBox9.Checked = true;
                checkBox10.Checked = true;
                checkBox11.Checked = true;
                checkBox12.Checked = true;
                checkBox13.Checked = true;
                checkBox14.Checked = true;
                checkBox15.Checked = true;
                checkBox16.Checked = true;

                checkBox1.Enabled = false;
                checkBox2.Enabled = false;
                checkBox3.Enabled = false;
                checkBox4.Enabled = false;
                checkBox5.Enabled = false;
                checkBox6.Enabled = false;
                checkBox7.Enabled = false;
                checkBox8.Enabled = false;
                checkBox9.Enabled = false;
                checkBox10.Enabled = false;
                checkBox11.Enabled = false;
                checkBox12.Enabled = false;
                checkBox13.Enabled = false;
                checkBox14.Enabled = false;
                checkBox15.Enabled = false;
                checkBox16.Enabled = false;


            }// end if
            else
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = false;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
                checkBox15.Checked = false;
                checkBox16.Checked = false;

                checkBox1.Enabled = true;
                checkBox2.Enabled = true;
                checkBox3.Enabled = true;
                checkBox4.Enabled = true;
                checkBox5.Enabled = true;
                checkBox6.Enabled = true;
                checkBox7.Enabled = true;
                checkBox8.Enabled = true;
                checkBox9.Enabled = true;
                checkBox10.Enabled = true;
                checkBox11.Enabled = true;
                checkBox12.Enabled = true;
                checkBox13.Enabled = true;
                checkBox14.Enabled = true;
                checkBox15.Enabled = true;
                checkBox16.Enabled = true;
            }// end else
        }  // end checkBox17_CheckedChanged
    }// end class

    public class NoteSet
    {
        public string server;
        public string port;
        public string user;
        public string pass;

        public string recipients;

        public bool highLev;
        public bool medLev;
        public bool allLev;

        public bool stat0;
        public bool stat1;
        public bool stat2;
        public bool stat3;
        public bool stat4;
        public bool stat5;
        public bool stat6;
        public bool stat7;
        public bool stat8;
        public bool stat9;
        public bool stat10;
        public bool stat11;
        public bool stat12;
        public bool stat13;
        public bool stat14;
        public bool stat15;

        public bool all;

        public bool on;
        public bool off;

    }

}
