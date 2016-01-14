using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

namespace NewBTASProto
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static Mutex mutex = new Mutex(true, "BTAS-16K");
        [STAThread]
        static void Main(String[] args)
        {
            if (mutex.WaitOne(TimeSpan.Zero, true))
            {
                try
                {
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
                MessageBox.Show("BTAS-16K is already running! Check the task bar.");
            }

        }
    }

}
