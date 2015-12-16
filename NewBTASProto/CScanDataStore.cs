using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NewBTASProto
{
    public class CScanDataStore
    {    

        //scalling constants
        const double KV = 0.002441;             //for voltages under 10V
        const double KB=  0.009764;             //for battery voltages
        const double KC = 0.0004882;            //for cell voltages
        const double KH = 0.004883;             //for voltages greater than 10V
        const double KS = 0.02441;              //for shunt readings
        const double KT = 0.0004882;            //'

        public int terminalID;
        public string programVersion;
        public int QS1;
        public byte TCAB;
        public string tempPlateType;
        private byte BCM;
        public byte CCID;
        public string cellCableType;
        public byte SHCID;
        public string shuntCableType;
        public int CTR;
        public string WDO;
        public double ch0GND;
        public double plus5V;
        public double ACMETempSen;
        public double currentOne;

        public double VB4;
        public double VB3;
        public double VB2;
        public double VB1;

        public double TP1;
        public double TP2;
        public double TP3;
        public double TP4;
        public double TP5;

        public double currentTwo;

        public double minus15;
        public double plus15;

        public double cellGND1;
        public double cellGND2;
        public double[] orderedCells = new double[24];
        private double[] cells = new double[24];

        public double ref95V;

        public int technology;

        public int customNoCells;
        public int batNumCable10;

        public int cellsToDisplay;

        // these are related to the status of the charger connected to the CSCAN
        // does not apply to ICs
        public bool cHold = false;
        public bool connected = false;
        public bool powerOn = false;
        public bool cycleEnd = false;

        // the constructor pulls in the data and stores it in the familiar A
        public CScanDataStore(string[] CDATA)
        {
            //terminalID = A[1]-216;
            terminalID = Int32.Parse(CDATA[1]) - 216;
 
            //stored in A[2];
            programVersion = CDATA[2];

            //this is a status field equals A[3] - 1000
            QS1 = Int32.Parse(CDATA[3]) - 1000;
            //connected?
            if (((byte) QS1 & 0X08) != 0) { connected = false; }
            else { connected = true; }
            //powerOn?
            if (((byte)QS1 & 0X04) != 0) { powerOn = false; }
            else { powerOn = true; }

            //TCAB represents the temp plate type
            TCAB = (byte) (Int32.Parse(CDATA[3]) - 1000);
            TCAB = (byte) (TCAB & 0x03);
            // TCB = 3 for none
            // TCB = 1 for TEMP-PLATE
            // TCB = 0 for TEST PLUG
            switch(TCAB){
                case 0:
                    tempPlateType = "TEST BOX";
                    break;
                case 1:
                    tempPlateType = "TEMP-PLATE";
                    break;
                case 3:
                    tempPlateType = "NONE";
                    break;
                default:
                    tempPlateType = "BAD VALUE";
                    break;
            }

            //BCM represents the cells cable being used
            //BCM = 255 - (A[4] - 1000)
            BCM = (byte)(255 - (Int32.Parse(CDATA[4]) - 1000));
            CCID = (byte)(BCM & 0x1F);
            switch (CCID)
            {
                case 0:
                    cellCableType = "NONE";
                    cellsToDisplay = 0;
                    break;
                case 1:
                    cellCableType = "20 CELLS";
                    cellsToDisplay = 20;
                    break;
                case 2:
                    cellCableType = "19 CELLS";
                    cellsToDisplay = 19;
                    break;
                case 10:
                    cellCableType = "4 BATT";
                    cellsToDisplay = 0;
                    break;
                case 31:
                    cellCableType = "CELL SIM";
                    cellsToDisplay = 24;
                    break;
                case 3:
                    cellCableType = "2X11 Cable";
                    cellsToDisplay = 22;
                    break;
                case 4:
                    cellCableType = "3X7 Cable";
                    cellsToDisplay = 21;
                    break;
                case 21:
                    cellCableType = "21 CELLS";
                    cellsToDisplay = 21;
                    break;
                default:
                    cellCableType = "Unknown Cable";
                    cellsToDisplay = 24;
                    break;
            }

            // Now we look at BCM again to come up with the Shunt cable attached to the CSCAN
            SHCID = (byte)((BCM & 0xE0) >> 5);
            switch (SHCID)
            {
                case 0:
                    shuntCableType = "NONE";
                    break;
                case 1:
                    shuntCableType = "100A";
                    break;
                case 2:
                    shuntCableType = "2A";
                    break;
                case 3:
                    shuntCableType = "20A";
                    break;
                case 4:
                    shuntCableType = "200mA";
                    break;
                case 5:
                    shuntCableType = "2A / 10A";
                    break;
                case 7:
                    shuntCableType = "TEST BOX";
                    break;
                default:
                    break;
            }

            // CTR = A[5] - 10000
            CTR = Int32.Parse(CDATA[5]) - 10000;

            //WDO = A[6]
            WDO = CDATA[6];

            //ch0GND = subQNEG(A[7]) * KV
            ch0GND = subQNEG(CDATA[7]) * KV;

            //+5V  = subQNEG(A[8]) * KV
            plus5V = subQNEG(CDATA[8]) * KV;

            //ACMEtemp = (A[9] - 1000) * KV
            ACMETempSen = Int32.Parse(CDATA[9]) - 1000 * KV;

            //Current #1 = subANEG(A[10]) and then needs to be scalled for the shunt cable being used
            currentOne = subQNEG(CDATA[10]);
            switch (SHCID)
            {
                case 1:
                case 7:
                    currentOne *= KS;
                    break;
                default:
                    currentOne *= KV;
                    break;
            }

            // VB4 = subQNEG(A[11]) * KB
            VB4 = subQNEG(CDATA[11]) * KB;
            // VB3 = subQNEG(A[12]) * KB
            VB3 = subQNEG(CDATA[12]) * KB;
            // VB2 = subQNEG(A[13]) * KB
            VB2 = subQNEG(CDATA[13]) * KB;
            // VB1 = subQNEG(A[14]) * KB
            VB1 = subQNEG(CDATA[14]) * KB;

            //TP1 = therLin(A[15] - 1000)
            TP1 = therLin(Int32.Parse(CDATA[15]) - 1000);
            //TP2 = therLin(A[16] - 1000)
            TP2 = therLin(Int32.Parse(CDATA[16]) - 1000);
            //TP3 = therLin(A[17] - 1000)
            TP3 = therLin(Int32.Parse(CDATA[17]) - 1000);
            //TP4 = therLin(A[18] - 1000)
            TP4 = therLin(Int32.Parse(CDATA[18]) - 1000);
            //TP5 = therLin(A[19] - 1000)
            TP5 = therLin(Int32.Parse(CDATA[19]) - 1000);

            // currentTwo = subQNEG(A[20]) with a decision swtich for scalling
            currentTwo = subQNEG(CDATA[20]);
            switch (SHCID)
            {
                case 0:
                    currentTwo *= KV;
                    break;
                case 2:
                case 3:
                case 4:
                    currentTwo *= KT;
                    break;
                case 5:
                    currentTwo *= KT;
                    break;
                case 7:
                    currentTwo *= (KS*2);
                    break;
            }

            // -15 = subQNEG(A[21]) * KH
            minus15 = subQNEG(CDATA[21]) * KH;
            // +15 = subQNEG(A[22]) * KH
            plus15 = subQNEG(CDATA[22]) * KH;

            // cells ground #1 = subQNEG(A[23]) * KC
            cellGND1 = subQNEG(CDATA[23]) * KC;

            //first group of cells
            for (int i = 0; i < 15; i++) { cells[i] = subQNEG(CDATA[i + 24]) * KC; }

            // cells ground #2 = subQNEG(A[39]) * KC
            cellGND2 = subQNEG(CDATA[39]) * KC;

            //second group of cells
            for (int i = 0; i < 9; i++) { cells[i + 15] = subQNEG(CDATA[i + 40]) * KC; }

            // and finally we look at the 9.5V refernece, ref95V = subQNEG(A[49]) * KV
            ref95V = subQNEG(CDATA[49]) * KV;

            //for the last step we need to produce an array with the cells in order...
            orderedCells[12] = cells[0];
            orderedCells[9] = cells[1];
            orderedCells[10] = cells[2];
            orderedCells[6] = cells[3];
            orderedCells[3] = cells[4];
            orderedCells[4] = cells[5];
            orderedCells[0] = cells[6];
            orderedCells[13] = cells[7];
            orderedCells[14] = cells[8];
            orderedCells[11] = cells[9];
            orderedCells[7] = cells[10];
            orderedCells[8] = cells[11];
            orderedCells[5] = cells[12];
            orderedCells[1] = cells[13];
            orderedCells[2] = cells[14];
            orderedCells[23] = cells[15];
            orderedCells[19] = cells[16];
            orderedCells[20] = cells[17];
            orderedCells[22] = cells[18];
            orderedCells[18] = cells[19];
            orderedCells[17] = cells[20];
            orderedCells[15] = cells[21];
            orderedCells[21] = cells[22];
            orderedCells[16] = cells[23];

            // and set the technology
            technology = 0;
            customNoCells = 0;
            batNumCable10 = 0;

           // TODO set this variable to be changed when cable 10 is used and the user designates that lead acid batteries are being tested


        }

        public CScanDataStore()
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
