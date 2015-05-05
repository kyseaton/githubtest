﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NewBTASProto
{
    public class ICDataStore
    {
        // 21 members to allow access to
        public int terminalID;          //terminal ID
        public int boardID;             //board version ID
        public int PV1;                 //main program version
        public int PV2;                 // comms program version
        public int CTR;                 // Sample Counter
        public char WDO;                // Work Command
        public int QS1;                 // System Status 1
        public int QS2;                 // System Status 2
        public float refVolt;           // reference voltage
        public float battCurrent;       // battery current
        public float battVoltage;       // battery voltage
        public int ACVoltage;         // AC voltage
        public float backupBattVolt;    // backup batttery voltage
        public float BT1;               // battery temperature 1
        public float BT2;               // battery temperature 2
        public float BT3;               // battery temperature 3
        public float BT4;               // battery temperature 4
        public float AmbientTemp;       // Ambient Temperature
        public float HSTemp1;           // Heat Sink #1 temperature
        public float HSTemp2;           // Heat Sink #2 temperature
        public int AuxIn;               // Auxiliary Input

        bool online = false;
        bool run = false;
        bool hold = false;

        // the constructor pulls in the data and stores it in the familiar A
        public ICDataStore(string[] ICDATA)
        {
            terminalID = int.Parse(ICDATA[1]) - 100;
            boardID = int.Parse(ICDATA[2]) - 1000;
            PV1 = int.Parse(ICDATA[3]) - 1000;
            PV2 = int.Parse(ICDATA[4]) - 1000;
            CTR = int.Parse(ICDATA[7]) - 1000;
            WDO = char.Parse(ICDATA[8]);
            QS1 = int.Parse(ICDATA[5]) - 1000;
            QS2 = int.Parse(ICDATA[6]) - 1000;
            refVolt = (float.Parse(ICDATA[9]) - 1000) / 1000;

            if (boardID < 2)
            {
                battCurrent = (float.Parse(ICDATA[10]) - 1000) / 1000;
            }
            else
            {
                battCurrent = (float.Parse(ICDATA[10]) - 1000);
            }

            battVoltage = (float.Parse(ICDATA[11]) - 1000) / 10;
            ACVoltage = int.Parse(ICDATA[12]) - 1000;
            backupBattVolt = (float.Parse(ICDATA[13]) - 1000) / 100;
            BT1 = (float.Parse(ICDATA[14]) - 1000) / 10;
            BT2 = (float.Parse(ICDATA[15]) - 1000) / 10;
            BT3 = (float.Parse(ICDATA[16]) - 1000) / 10;
            BT4 = (float.Parse(ICDATA[17]) - 1000) / 10;
            AmbientTemp = (float.Parse(ICDATA[18]) - 1000) / 10;
            HSTemp1 = (float.Parse(ICDATA[19]) - 1000) / 4;
            HSTemp2 = (float.Parse(ICDATA[20]) - 1000) / 4;
            AuxIn = int.Parse(ICDATA[21]) - 1000;
            
            if((QS1 & 0x01) == 1)
            {
                online = true;
            }
            if ((QS1 & 0x02) == 2)
            {
                run = true;
            }
            if ((QS1 & 0x04) == 4)
            {
                hold = true;
            }

        }

        public ICDataStore()
        {
            // TODO: Complete member initialization
        }

        // got this from the original VB6 program
        private int subQNEG(string RVAL)
        {
            int temp = Int32.Parse(RVAL);
            if (temp < 1000 || temp > 9191) { return 1000; }
            else
            {
                temp = temp - 1000;
                if (temp > 4096)
                {
                    temp = temp - 8192;
                }
                return temp;
            }
        }

        //also got this from the original VB6 program
        // this is the segmented thermistor linearization routine
        private double therLin(int THV)
        {
            int temp = THV;

            if (temp > 4080) { return -99; }
            else if (temp > 3624) { return -98; }
            else if (temp < 10) { return -96; }
            else if (temp < 820) { return -97; }
            else if (temp <= 941) { return (89 + 0.0496 * (941 - temp)); }
            else if (temp <= 1079) { return (83 + 0.0436 * (1079 - temp)); }
            else if (temp <= 1234) { return (77 + 0.0386 * (1234 - temp)); }
            else if (temp <= 1408) { return (71 + 0.0346 * (1408 - temp)); }
            else if (temp <= 1598) { return (65 + 0.0314 * (1598 - temp)); }
            else if (temp <= 1805) { return (59 + 0.029 * (1805 - temp)); }
            else if (temp <= 2025) { return (53 + 0.0273 * (2025 - temp)); }
            else if (temp <= 2253) { return (47 + 0.0263 * (2253 - temp)); }
            else if (temp <= 2485) { return (41 + 0.0259 * (2485 - temp)); }
            else if (temp <= 2714) { return (35 + 0.0262 * (2714 - temp)); }
            else if (temp <= 2934) { return (29 + 0.0273 * (2934 - temp)); }
            else if (temp <= 3138) { return (23 + 0.0293 * (3138 - temp)); }
            else if (temp <= 3323) { return (17 + 0.0324 * (3323 - temp)); }
            else if (temp <= 3486) { return (11 + 0.037 * (3486 - temp)); }
            else if (temp <= 3624) { return (5 + 0.0434 * (3624 - temp)); }
            else { return -98; }

        }
    }
}
