using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NewBTASProto
{
    public class ICSettingStore
    {
        //OUTPUT DATA STRUCTURE ----------------------------------
        public byte T1;
        public byte T2;

        public byte WDO;

        //--- P ---
        public byte KE1;                                           
        //KE1 is defined at the beginning of this subroutine: 0=query, 1=data, 2=command, 3=data/command
        //query: read status and data only
        //send Battery Test Profile programming parameters
        //command: issue a command (start,stop, reset)
        //data/command: program parameters and issue a command
        public byte KE2;                                           //KE2 not used for now (3 bits, available)
        public byte KE3;                                           //action: 0=clear, 1=run, 2=stop, 3=reset
        public byte KM0;    //first command [type + test + action] lower 8 bits   KM0 = (byte)((KE1 + 4 * KE2 + 64 * KE3) + 48)
        public byte KM1;                            //'Mode (10, 11, 12, 21, 22, 30, 31, 32 [for now] )  KM1 = (byte)(10 + 48)

        //--- A ---
        public byte KM2;                             //CT1H, Charge Time 1, Hours
        public byte KM3;                             //CT1M, Charge Time 1, Minutes
        public byte KM4;                             //CC1H, Charge current 1, High (byte)
        public byte KM5;                             //CC1L, Charge current 1, Low (byte)
        public byte KM6;                             //CV1H, Charge voltage 1, High (byte)
        public byte KM7;                             //CV1L, Charge voltage 1, Low (byte)

        //--- B ---
        public byte KM8;                             //CT2H, Charge Time 2, Hours
        public byte KM9;                             //CT2M, Charge Time 2, Minutes
        public byte KM10;                             //CC2H, Charge Current 2, High (byte)
        public byte KM11;                            //CC2L, Charge Current 2, Low (byte)
        public byte KM12;                            //CV2H, Charge Voltage 2, High (byte)
        public byte KM13;                            //CV2L, Charge voltage 2, Low (byte)

        //--- C ---
        public byte KM14;                            //DTH, Discharge Time, Hours
        public byte KM15;                             //DTM, Discharge Time, Minutes
        public byte KM16;                            //DCH, Discharge current, High (byte)
        public byte KM17;                            //DCL, Discharge current, Low (byte)
        public byte KM18;                            //DVH, Discharge voltage, High (byte)
        public byte KM19;                            //DVL, Discharge Voltage, Low (byte)
        public byte KM20;                            //CRH, Discharge Resistance, High (byte)
        public byte KM21;                            //CRL, Discharge Resistance, Low (byte)

        public byte ULCH;                                          //---

        public byte[] outText = new byte[28];  // and we'll put this here to make the IC code a little cleaner

        public ICSettingStore(int j)
        {
            //OUTPUT DATA STRUCTURE ----------------------------------
            T1 = Convert.ToByte((j / 10 + 48).ToString("0"));
            T2 = Convert.ToByte((j % 10 + 48).ToString("0"));

            WDO = (byte)'L';

            //--- P ---
            KE1 = 0;                                           //KE1 is defined at the beginning of this subroutine: 0=query, 1=data, 2=command, 3=data/command
            //query: read status and data only
            //send Battery Test Profile programming parameters
            //command: issue a command (start,stop, reset)
            //data/command: program parameters and issue a command
            KE2 = 0;                                           //KE2 not used for now (3 bits, available)
            KE3 = 3;                                           //action: 0=clear, 1=run, 2=stop, 3=reset
            KM0 = (byte)((KE1 + 4 * KE2 + 64 * KE3) + 48);    //first command [type + test + action] lower 8 bits
            KM1 = (byte)(10 + 48);                            //'Mode (10, 11, 12, 21, 21, 30, 31, 32 [for now] )
            
            //--- A ---
            KM2 = (byte)'0';                             //CT1H, Charge Time 1, Hours
            KM3 = (byte)'0';                             //CT1M, Charge Time 1, Minutes
            KM4 = (byte)'0';                             //CC1H, Charge current 1, High (byte)
            KM5 = (byte)'0';                             //CC1L, Charge current 1, Low (byte)
            KM6 = (byte)'0';                             //CV1H, Charge voltage 1, High (byte)
            KM7 = (byte)'0';                             //CV1L, Charge voltage 1, Low (byte)

            //--- B ---
            KM8 = (byte)'0';                             //CT2H, Charge Time 2, Hours
            KM9 = (byte)'0';                             //CT2M, Charge Time 2, Minutes
            KM10 = (byte)'0';                             //CC2H, Charge Current 2, High (byte)
            KM11 = (byte)'0';                            //CC2L, Charge Current 2, Low (byte)
            KM12 = (byte)'0';                            //CV2H, Charge Voltage 2, High (byte)
            KM13 = (byte)'0';                            //CV2L, Charge voltage 2, Low (byte)

            //--- C ---
            KM14 = (byte)'0';                            //DTH, Discharge Time, Hours
            KM15 = (byte)'0';                             //DTM, Discharge Time, Minutes
            KM16 = (byte)'0';                            //DCH, Discharge current, High (byte)
            KM17 = (byte)'0';                            //DCL, Discharge current, Low (byte)
            KM18 = (byte)'0';                            //DVH, Discharge voltage, High (byte)
            KM19 = (byte)'0';                            //DVL, Discharge Voltage, Low (byte)
            KM20 = (byte)'0';                            //CRH, Discharge Resistance, High (byte)
            KM21 = (byte)'0';                            //CRL, Discharge Resistance, Low (byte)

            ULCH = (byte)'0';                                          //---

            // and lastly, we'll fill in the output array
            outText[0] = (byte)'~';
            outText[1] = T1;
            outText[2] = T2;
            outText[3] = WDO;
            outText[4] = KM0;
            outText[5] = KM1;
            outText[6] = KM2;
            outText[7] = KM3;
            outText[8] = KM4;
            outText[9] = KM5;
            outText[10] = KM6;
            outText[11] = KM7;
            outText[12] = KM8;
            outText[13] = KM9;
            outText[14] = KM10;
            outText[15] = KM11;
            outText[16] = KM12;
            outText[17] = KM13;
            outText[18] = KM14;
            outText[19] = KM15;
            outText[20] = KM16;
            outText[21] = KM17;
            outText[22] = KM18;
            outText[23] = KM19;
            outText[24] = KM20;
            outText[25] = KM21;
            outText[26] = ULCH;
            outText[27] = (byte)'Z'; //send wake-up character, terminal ID, WDO, commands and Z
        }

        public void UpdateOutText(){
            outText[0] = (byte)'~';
            outText[1] = T1;
            outText[2] = T2;
            outText[3] = WDO;
            KM0 = (byte)((KE1 + 4 * KE2 + 64 * KE3) + 48);
            outText[4] = KM0;
            outText[5] = KM1;
            outText[6] = KM2;
            outText[7] = KM3;
            outText[8] = KM4;
            outText[9] = KM5;
            outText[10] = KM6;
            outText[11] = KM7;
            outText[12] = KM8;
            outText[13] = KM9;
            outText[14] = KM10;
            outText[15] = KM11;
            outText[16] = KM12;
            outText[17] = KM13;
            outText[18] = KM14;
            outText[19] = KM15;
            outText[20] = KM16;
            outText[21] = KM17;
            outText[22] = KM18;
            outText[23] = KM19;
            outText[24] = KM20;
            outText[25] = KM21;
            outText[26] = ULCH;
            outText[27] = (byte)'Z'; //send wake-up character, terminal ID, WDO, commands and Z
        }

    }
}
