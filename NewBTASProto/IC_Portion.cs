using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;

namespace NewBTASProto
{
    public partial class Main_Form : Form
        {
        
        //ComPort.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(port_DataReceived_1);
        //EventArgs E = new EventArgs();

        /// <summary>
        /// Serial Stuff defined here
        /// </summary>
        public SerialPort ICComPort = new SerialPort();

        public CancellationTokenSource cPollIC;

        //for checking the IC when selected
        bool check =  false;
        int toCheck;
        int chanNum;

        //for critical operations (Start,Stop, etc)
        bool [] criticalNum = new bool[16];

        //Com Error count 
        byte[] comErrorNum = new byte[16] {0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0};

        public void pollICs()
        {

            //This code is here to poll the ICs/////////////////////////////////////////////////////////////////////////////////////
            // We are going to run all of this code on a helper thread, as to improve GUI performance//////////////////////////////////////////

            cPollIC = new CancellationTokenSource();

            string tempBuff = "";
            ICDataStore testData;

            // Open the comport
            ICComPort.ReadTimeout = 100;
            ICComPort.PortName = GlobalVars.ICComPort;
            ICComPort.BaudRate = 19200;
            ICComPort.DataBits = 8;
            ICComPort.StopBits = StopBits.One;
            ICComPort.Handshake = Handshake.None;
            ICComPort.Parity = Parity.None;
            ICComPort.DtrEnable = true;
            ICComPort.RtsEnable = false;
            ICComPort.Open();

            //OUTPUT DATA STRUCTURE ----------------------------------

            ThreadPool.QueueUserWorkItem(s =>
            {
                CancellationToken token = (CancellationToken)s;
                Thread.Sleep(1500);

                while (true)
                {
                    
                    try
                    {
                        for (int j = 0; j < 16; j++)
                        {
                            //sleep a little each time as to not overload the host
                            Thread.Sleep(10);
                            //putting the cancellation token in a often looked at place...
                            if (token.IsCancellationRequested) return;

                            ////////////////////////////////////////////NORMAL PRIORITY CHARGERS ARE CHECKED HERE///////////////////////
                            if ((bool) d.Rows[j][8] && (bool) d.Rows[j][4] && (string) d.Rows[j][9] != "" && d.Rows[j][10].ToString().Contains("ICA"))
                            {
                                Thread.Sleep(600);
                                try
                                {
                                    // control for master slave setup
                                    int chargerID = 0;

                                    if (d.Rows[j][9].ToString() == "") { ;}  // do nothing if there is no assigned charger id
                                    else if (d.Rows[j][9].ToString().Length > 2)  // this is the case where we have a master and slave config
                                    {
                                        // we have a master slave charger
                                        // split into 3 and 4 digit case
                                        if (d.Rows[j][9].ToString().Length == 3)
                                        {
                                            if (d.Rows[j][9].ToString().Substring(2, 1) == "S") { break; }
                                            // 3 case
                                            chargerID = int.Parse(d.Rows[j][9].ToString().Substring(0, 1));
                                        }
                                        else
                                        {
                                            if (d.Rows[j][9].ToString().Substring(3, 1) == "S") { break; }
                                            // 4 case
                                            chargerID = int.Parse(d.Rows[j][9].ToString().Substring(0, 2));

                                        }
                                    }
                                    else  // this is the normal case with just one charger
                                    {
                                        chargerID = Convert.ToInt32(d.Rows[j][9]);
                                    }


                                    // send the short command based on the settings for the charger...
                                    ICComPort.Write(GlobalVars.ICSettings[chargerID].outText, 0, 28);
                                    // wait for a response
                                    tempBuff = ICComPort.ReadTo("Z");
                                    //do something with the new data
                                    char[] delims = { ' ' };
                                    string[] A = tempBuff.Split(delims);
                                    //A[1] has the terminal ID in it
                                    testData = new ICDataStore(A);

                                    //put this new data in the chart...
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        if (testData.online == true)
                                        {
                                            if (testData.faultStatus != "") { updateD(j, 11, testData.faultStatus); }
                                            else if (testData.endStatus != "") { updateD(j, 11, testData.endStatus);}
                                            else { updateD(j, 11, testData.runStatus); }
                                            if ((bool)d.Rows[j][8]) { dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Green; }
                                        }
                                        else 
                                        {
                                            updateD(j,11,"offline!");
                                            if ((bool)d.Rows[j][8]) { dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Red; }
                                        }

                                        // also update the type of charger being used
                                        if (testData.boardID == 1) { updateD(j, 10, "ICA mini"); }
                                        else if (testData.boardID == 6) { updateD(j, 10, "ICA SMC"); }
                                        else if (testData.boardID == 8) { updateD(j, 10, "ICA SMC ED"); }
                                        else if (testData.boardID == 6) { updateD(j, 10, "ICA SMini"); }
                                        
                                        rtbIncoming.Text = j.ToString() + "  :  " + tempBuff;
                                        
                                    });

                                    comErrorNum[j] = 0;

                                    Thread.Sleep(200);

                                }
                                catch (Exception ex)
                                {
                                    if (ex is System.TimeoutException)
                                    {
                                        if (comErrorNum[j] < 3) { comErrorNum[j]++; }
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            if (comErrorNum[j] > 2)
                                            {
                                                updateD(j, 11, "");
                                            }
                                            tempBuff = ICComPort.ReadExisting();
                                            rtbIncoming.Text = "Com Error" + System.Environment.NewLine + tempBuff;
                                        });
                                        Thread.Sleep(100);
                                    }
                                    else { throw ex; }
                                }       // end catch
                            }       // end if
                                
                            else if(d.Rows[j][10].ToString().Contains("ICA")) 
                            {
                                if ((string)d.Rows[j][11] != "")
                                {
                                    updateD(j, 11, "");
                                }
                                dataGridView1.Rows[j].Cells[8].Style.BackColor = Color.Gainsboro;
                              
                            }

                            /////////////////////////////CHECK FOR CHARGER IDENTITY///////////////////////////////////////
                            if (check)
                            {
                                Thread.Sleep(500);
                                try
                                {
                                    // send the short command based on the settings for the charger...
                                    ICComPort.Write(GlobalVars.ICSettings[toCheck].outText, 0, 28);
                                    // wait for a response
                                    tempBuff = ICComPort.ReadTo("Z");
                                    //do something with the new data
                                    char[] delims = { ' ' };
                                    string[] A = tempBuff.Split(delims);
                                    //A[1] has the terminal ID in it
                                    testData = new ICDataStore(A);
                                    // if we got one then we can determine that we have an ICA
                                    if (testData.boardID == 1) { updateD(chanNum, 10, "ICA mini"); }
                                    else if (testData.boardID == 6) { updateD(chanNum, 10, "ICA SMC"); }
                                    else if (testData.boardID == 8) { updateD(chanNum, 10, "ICA SMC ED"); }
                                    else if (testData.boardID == 6) { updateD(chanNum, 10, "ICA SMini"); }
                                    // and we don't need to check any more
                                    check = false;
                                    Thread.Sleep(200);
                                }
                                catch (Exception ex)
                                {
                                    if (ex is System.TimeoutException)
                                    {
                                        Thread.Sleep(100);
                                    }
                                    else { throw ex; }
                                }       // end catch
                            }       // end else if

                            ////////////////////////////////////////////HIGH PRIORTIY CHARGERS ARE CHECKED HERE///////////////////////
                            // we need to check for critical operation also!
                            for(int i = 0;i < 16; i++)
                            {
                                if (criticalNum[i] == true)
                                {
                                    try
                                    {
                                        Thread.Sleep(500);
                                        // send the short command based on the settings for the charger...
                                        ICComPort.Write(GlobalVars.ICSettings[i].outText, 0, 28);
                                        // wait for a response
                                        tempBuff = ICComPort.ReadTo("Z");

                                        int station = 0;
                                        //find where the charger is located
                                        for (byte v = 0; v < 16; v++)
                                        {
                                            if (d.Rows[v][9].ToString() == "") { ;}  // do nothing if there is no assigned charger id
                                            else if (d.Rows[v][9].ToString().Length > 2)  // this is the case where we have a master and slave config
                                            {
                                                // we have a master slave charger
                                                // split into 3 and 4 digit case
                                                if (d.Rows[v][9].ToString().Length == 3)
                                                {
                                                    // 3 case
                                                    station = int.Parse(d.Rows[v][9].ToString().Substring(0,1));
                                                }
                                                else
                                                {
                                                    // 4 case
                                                    station = int.Parse(d.Rows[v][9].ToString().Substring(0, 2));

                                                }
                                            }
                                            else if (int.Parse(d.Rows[v][9].ToString()) == i)  // this is the normal case of just one charger
                                            {
                                                station = v;
                                                break;
                                            }
                                        }


                                        // we got a response so lets update the grid and the status box
                                        //A[1] has the terminal ID in it
                                        char[] delims = { ' ' };
                                        string[] A = tempBuff.Split(delims);
                                        testData = new ICDataStore(A);
                                        this.Invoke((MethodInvoker)delegate
                                        {
                                            rtbIncoming.Text = "Critical  " + i.ToString() + "  :  " + tempBuff;
                                            if (testData.online == true)
                                            {
                                                if (testData.faultStatus != "") { updateD(station, 11, testData.faultStatus); }
                                                else if (testData.endStatus != "") { updateD(station, 11, testData.endStatus); }
                                                else { updateD(station, 11, testData.runStatus);}
                                                if ((bool)d.Rows[station][8]) { dataGridView1.Rows[station].Cells[8].Style.BackColor = Color.Green; }
                                            }
                                            else 
                                            {
                                                updateD(station, 11, "offline!");
                                                if ((bool)d.Rows[station][8]) { dataGridView1.Rows[station].Cells[8].Style.BackColor = Color.Red; }
                            
                                            }

                                            // also update the type of charger being used
                                            if (testData.boardID == 1) { updateD(station, 10, "ICA mini"); }
                                            else if (testData.boardID == 6) { updateD(station, 10, "ICA SMC"); }
                                            else if (testData.boardID == 8) { updateD(station, 10, "ICA SMC ED"); }
                                            else if (testData.boardID == 6) { updateD(station, 10, "ICA SMini"); }

                                        });
                                        Thread.Sleep(200);
                                    }
                                    catch (Exception ex)
                                    {
                                        if (ex is System.TimeoutException)
                                        {
                                            Thread.Sleep(100);
                                        }
                                        else { throw ex; }
                                    }       // end catch
                                } // end if
                            }// end for

                        }           // end for
                    }               // end try
                    catch (Exception ex)
                    {
                        if (token.IsCancellationRequested) return;
                        else
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }

                }                   // end while
            },cPollIC.Token); // end thread


        }  // end pollICs

        private void checkForIC(int chargerNum, int channelNum)
        {
            check = true;
            toCheck = chargerNum;
            chanNum = channelNum;
            // create a thread to check if we've got an IC connected at the selected address
             ThreadPool.QueueUserWorkItem(s =>
            {
                // we're going to give it 3 seconds to think about it...
                // it gets checked every other time...
                Thread.Sleep(12000);
                // now we'll make sure we're not looking anymore...
                check = false;
            }); // end thread


        }


    }
}
