
namespace NewBTASProto
{
    public static class GlobalVars
    {
        /// <summary>
        /// Indiates whether to use Fahrenheit (True) or Celsius (False)
        /// </summary>
        public static bool useF;

        /// <summary>
        /// Indiates whether to use Pos2Neg cell ordering (True) or Negative to Positive ordering (False)
        /// </summary>
        public static bool Pos2Neg;

        /// <summary>
        /// Stores the Buisness name (Used in Reports, etc...)
        /// </summary>
        public static string businessName;

        /// <summary>
        /// Selects where ther the highlight feature is on or off
        /// </summary>
        public static bool highlightCurrent;

        /// <summary>
        /// Stores the comport used for the CScans
        /// </summary>
        public static string CSCANComPort;

        /// <summary>
        /// Stores the comport used for the IC
        /// </summary>
        public static string ICComPort;

        /// <summary>
        /// Stores the settings for all possible attached Chargers
        /// </summary>
        public static ICSettingStore[] ICSettings =  new ICSettingStore[16];

        /// <summary>
        /// Stores the current data for all possible attached Chargers
        /// </summary>
        public static ICDataStore[] ICData = new ICDataStore[16];

        /// <summary>
        /// Stores the current data for all the attached CScans
        /// </summary>
        public static CScanDataStore[] CScanData = new CScanDataStore[16];

        /// <summary>
        /// Indicates if we are going to automatically configure the ICs
        /// </summary>
        public static bool autoConfig;

        /// <summary>
        /// Stores the Current Technician
        /// </summary>
        public static string currentTech;

        /// <summary>
        /// this is the loading bool
        /// </summary>
        public static bool loading = true;

        /// <summary>
        /// this array will determine if the cscan is holding the charger
        /// this bit needs to be cleared when the test is going to be run
        /// </summary>
        public static bool[] cHold = new bool[16] {
            false,false,false,false,
            false,false,false,false,
            false,false,false,false,
            false,false,false,false};

        /// <summary>
        /// this array will determine if Current #2 value should be displayed
        /// </summary>
        public static bool[] curr2Dis = new bool[16] {
            false,false,false,false,
            false,false,false,false,
            false,false,false,false,
            false,false,false,false};

        /// <summary>
        /// this string holds the programversion text
        /// </summary>
        public static string programVersion = "6.0.1";

        /// <summary>
        /// this string holds the program publish date
        /// </summary>
        public static string programPubDate = "Aug 09 2016";

        /// <summary>
        /// this is where we hold our notification service settings...
        /// </summary>
        public static string server;
        public static string port;
        public static string user;
        public static string pass;

        public static string recipients;

        public static bool highLev;
        public static bool medLev;
        public static bool allLev;

        public static bool stat0;
        public static bool stat1;
        public static bool stat2;
        public static bool stat3;
        public static bool stat4;
        public static bool stat5;
        public static bool stat6;
        public static bool stat7;
        public static bool stat8;
        public static bool stat9;
        public static bool stat10;
        public static bool stat11;
        public static bool stat12;
        public static bool stat13;
        public static bool stat14;
        public static bool stat15;

        public static bool all;

        public static bool noteOn;
        public static bool noteOff;

        /// <summary>
        /// masterFiller Vars
        /// </summary>
        public static bool checkMasterFiller = false;

        public static string[] MFData = new string[30];

        // Here we are going to store the saved setting values
        public static string BTAS16NVConnectionString;
        public static int FormWidth;
        public static int FormHeight;
        public static int PositionX;
        public static int PositionY;
        public static bool maximized;
        public static bool showSels;
        public static bool dualPlots;
        public static bool cb1;
        public static bool cb2;
        public static bool cb3;
        public static bool cb4;
        public static bool cb5;
        public static bool cb6;
        public static bool FC6C1MinimumCellVotageAfterChargeTestEnabled;
        public static decimal FC6C1MinimumCellVoltageThreshold;
        public static bool DecliningCellVoltageTestEnabled;
        public static bool FC6C1WaitEnabled;
        public static decimal FC6C1WaitTime;
        public static bool cbComplete;
        public static bool cbUpdateCompleteDate;
        public static bool FC4C1MinimumCellVotageAfterChargeTestEnabled;
        public static decimal FC4C1MinimumCellVoltageThreshold;
        public static bool FC4C1WaitEnabled;
        public static decimal FC4C1WaitTime;
        public static bool CapTestVarEnable;
        public static decimal CapTestVarValue;
        public static decimal CSErr2Allow;

    }
}
