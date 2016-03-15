using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.ExceptionServices;


namespace NewBTASProto
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static Mutex mutex = new Mutex(true, "BTAS-16K");
        [STAThread]
        [HandleProcessCorruptedStateExceptions]
        static void Main(String[] args)
        {

            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += new UnhandledExceptionEventHandler(MyHandler);
            //Application.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            if (mutex.WaitOne(TimeSpan.Zero, true))
            {
                try
                {
                    Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new Splash());
                }
                catch (Exception ex)
                {
                    MessageBox.Show("From Program.cs:  " + ex.Message + Environment.NewLine + ex.StackTrace);
                }
            }
            else
            {
                MessageBox.Show("BTAS-16K is already running! Check the task bar.", "Check Task Bar!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        static void MyHandler(object sender, UnhandledExceptionEventArgs args)
        {
            Exception e = (Exception)args.ExceptionObject;
            if (e is ThreadAbortException)
            {
                //do nothing
            }
            else
            {
                MessageBox.Show("Unhandled exception : " + e.Message + Environment.NewLine + e.StackTrace);
            }
            
            
            
        }

    }

}
