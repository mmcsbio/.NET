using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using System.IO;

namespace BayantechAddIn
{
    public partial class ThisAddIn
    {
        string licensePath = @"C:\Program Files\Common Files\win\word2010License.txt";
        public bool EXPIRED
        {
            set
            {
            }
            get
            {
                if (checkExpiration() == 0)
                    return true;
                else
                    return false;
            }
        }
        DateTime curDate = DateTime.Now;
        DateTime expDate = new DateTime(2016, 10, 1);
        //Number of days left to notify at
        int notificationDays = 7;
        //Number of times this AddIn is activated (for runtimes file counter)
        int activationTimes = 2;
        //Maximum number of times this AddIn can be run/month
        int runTimes = 120;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Check Expiration & Notification Block
            int daysLeft = checkExpiration();
            if (daysLeft == 0)
                EXPIRED = true;
            if (!EXPIRED)
            {
                if (daysLeft <= notificationDays)
                    showExpirationNotice(daysLeft);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        /// <summary>
        /// Check if the license file is created and contains data or not, create & initialize it if not
        /// </summary>
        private void checkLicense()
        {
            string dir = getLicenseDirectory();
            if(!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            StreamWriter writer = new StreamWriter(licensePath, true);
            if (!File.Exists(licensePath))
                writer.WriteLine("1");
            writer.Close();
            hideFile();
        }

        private void hideFile()
        {
            File.SetAttributes(licensePath, FileAttributes.Hidden);
        }

        private void showFile()
        {
            File.SetAttributes(licensePath, FileAttributes.Normal);
        }

        private string getLicenseDirectory()
        {
            //extract the directory of the license file
            List<string> dirSplit = licensePath.Split('\\').ToList();
            dirSplit.RemoveAt(dirSplit.Count - 1);
            return string.Join("\\", dirSplit);
        }
        /// <summary>
        ///Reads the "lincense file" lines and adds extra lines if it's length is less than "activationTimes variable"
        /// </summary>
        /// <returns>List of all file lines</returns>
        private List<string> readLicense()
        {
            checkLicense();
            List<string> lines = File.ReadAllLines(licensePath).ToList();
            List<string> newLines = new List<string>();
            while (lines.Count + newLines.Count < activationTimes)
            {
                //add the maximum runTimes in the extra lines(so that no one can run previous activations)
                newLines.Add((runTimes + 1).ToString());
            }

            if (newLines.Count > 0)
            {

                //reset the last line to the default value
                newLines[newLines.Count - 1] = "1";

                lines.AddRange(newLines);
                //Append the new lines to the end of the file
                File.AppendAllLines(licensePath, newLines);
            }
            return lines;
        }

        /// <summary>
        /// Write list of lines in license file with overriding existing data
        /// </summary>
        /// <param name="lines"></param>
        private void writeLicense(List<string> lines)
        {
            checkLicense();
            showFile();
            File.WriteAllLines(licensePath, lines);
            hideFile();
        }

        /// <summary>
        /// Increment the license file based on the activatioTimes number.
        /// </summary>
        public void incrementLicense()
        {
            int curRunTimes = 0;
            int incrementValue = 1;
            List<string> lines = readLicense();

            //incremented value
            curRunTimes = Int32.Parse(lines[activationTimes - 1]);

            //check the limits before incrementing
            if (curRunTimes <= runTimes)
            {
                curRunTimes += incrementValue;
                lines[activationTimes - 1] = curRunTimes.ToString();
            }

            writeLicense(lines);
        }
        /// <summary>
        /// Check days left for the AddIn to expire or if it is Expired
        /// </summary>
        /// <param name="curDate"></param>
        /// <param name="expDate"></param>
        /// <param name="activationTimes"></param>
        /// <returns>Number of Days left to expire or 0 if expired</returns>
        private int checkExpiration()
        {
            int daysLeft = (int)Math.Ceiling(expDate.Subtract(curDate).TotalDays);
            int curRunTimes = getRunTimes();
            if (daysLeft <= 0 || curRunTimes > runTimes)
                return 0;

            return daysLeft;
        }

        /// <summary>
        /// Get the number of times this AddIn has worked at a specified activation time
        /// </summary>
        /// <returns>The number of times the AddIn has worked</returns>
        private int getRunTimes()
        {
            int curRunTimes = 0;
            //activationTimes variable is 1 based numbered values
            string line = readLicense()[activationTimes - 1];
            //If can't parse the line then it might been manipulated, then set the curTimes exceeds the runTimes variable
            if (!Int32.TryParse(line, out curRunTimes))
                curRunTimes = runTimes + 1;

            return curRunTimes;
        }

        public void showExpired()
        {
            MessageBox.Show("Bayantech Addins has Expired!\nPlease contact your Manager or the R&D department to renew your license.", "Reactivate Bayantech AddIns License", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void showExpirationNotice(int daysLeft)
        {
            MessageBox.Show("Bayantech Addins will expire after " + daysLeft + " days!\nPlease contact your manager or R&D department to renew your license.", "Reactivate Bayantech AddIns License", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
