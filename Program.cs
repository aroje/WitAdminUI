using Microsoft.TeamFoundation.Client;
using System;
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

    public class Helpers
    {
        public static TfsTeamProjectCollection ConnectToTeamProject()
        {
            try
            {
                TeamProjectPicker tpp = new TeamProjectPicker();
                if (tpp.ShowDialog() == DialogResult.OK)
                {
                    return tpp.SelectedTeamProjectCollection;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            return null;
        }
    }
}
