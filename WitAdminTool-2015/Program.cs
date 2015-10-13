using Microsoft.TeamFoundation.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace WitAdminTool
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            TfsTeamProjectCollection tfsTPC = WitAdminTool.Helpers.ConnectToTeamProject();
            if (tfsTPC != null)
                Application.Run(new WitAdminToolForm(tfsTPC));
        }
    }
}
