using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using System.Windows.Forms;

namespace CloseToTray
{
    public partial class ThisAddIn
    {
        private OutlookWin32Window explrWin32Window;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                var explorer = Application.ActiveExplorer();
                if (explorer != null)
                    InitOutlookWin32Window(explorer);
                else
                    Application.Explorers.NewExplorer += Explorers_NewExplorer;

                if (EnsureMinToTrayIsSet())
                    MessageBox.Show("Setting 'Minimize to tray' changed to true.\nPlease restart outlook", "Settings changed", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error on initializing 'CloseToTray':\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Explorers_NewExplorer(Outlook.Explorer Explorer)
        {
            InitOutlookWin32Window(Explorer);
            Application.Explorers.NewExplorer -= Explorers_NewExplorer;
        }

        private void InitOutlookWin32Window(Outlook.Explorer explorer)
        {
            if(explorer == null || explrWin32Window != null)
                return;

            explrWin32Window = new OutlookWin32Window(explorer);
            explrWin32Window.Closing += ExplrWin32Window_Closing;
        }

        private void ExplrWin32Window_Closing(object sender, OutlookWin32Window.CancelEventArgs e)
        {
            e.Cancel = true;
            Application.ActiveExplorer().WindowState = Outlook.OlWindowState.olMinimized;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            explrWin32Window.Dispose();
        }

        private string GetOutlookVersion()
        {
            string[] versionParts = Application.Version.Split('.');
            return $"{versionParts[0]}.{versionParts[1]}";
        }

        private bool EnsureMinToTrayIsSet()
        {
            string outlookVersion = GetOutlookVersion();
            using (RegistryKey mttKey = Registry.CurrentUser.OpenSubKey($"SOFTWARE\\Microsoft\\Office\\{outlookVersion}\\Outlook\\Preferences", true))
            {
                if ((int?)mttKey.GetValue("MinToTray") != 1)
                {
                    mttKey.SetValue("MinToTray", 1);
                    return true;
                }
            }

            return false;
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
