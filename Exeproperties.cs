using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Tools.WindowsInstallerXml.Msi;
using Microsoft.Tools.WindowsInstallerXml;

namespace Exeproperties
{
    public class ExePropGenerator
    {
        private string strExeFilePath { get; set; }
        public Dictionary<string,string> InstallerProps { get; private set; }
        private static string tempFolder = Path.GetTempPath() + "ExePropGen" + DateTime.Now.ToString("yyyyMMddHHmmssfff");
        
        /// <summary>
        /// Constructor. Takes EXE file path as an input
        /// </summary>
        /// <param name="strInExePath"></param>
        public ExePropGenerator(string strInExePath)
        {
            strExeFilePath = strInExePath;
            InstallerProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Gets the properties for the given EXE
        /// </summary>
        /// <returns>true, if successful. false, if not successful</returns>
        public bool SetProperties()
        {
            bool isDone = false;

            // Create temp folder
            Directory.CreateDirectory(tempFolder);

            // Extract MSI to temp folder
            if (strExeFilePath.Substring(strExeFilePath.Length - 3, 3).ToUpper() != "MSI")
            {
                if (!ExtractMSI(tempFolder))
                    return isDone;
                // Get properties from MSI and update the InstallerProperties
                isDone = GetMSIProperties(GetMSIFile(tempFolder));
            }
            else
            {
                //if the input file itself is msi
                isDone = GetMSIProperties(strExeFilePath);
            }
            // Delete the temp folder
            try
            {
                Directory.Delete(tempFolder, true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            // Set boolean value based on success or failure of the entire flow
            return isDone;
        }

        private bool DeleteDirectory(string tempFolder)
        {
            bool isDone = false;

            DirectoryInfo dirInfo = new DirectoryInfo(tempFolder);
            try
            {
                foreach (FileInfo file in dirInfo.GetFiles())
                {
                    if (file.IsReadOnly)
                        file.IsReadOnly = false;

                    file.Delete();
                }

                foreach (DirectoryInfo dir in dirInfo.GetDirectories())
                {
                    dir.Delete(true);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return isDone;
            }
            isDone = true;
            return isDone;
        }

        /// <summary>
        /// Extracts MSI file from the EXE
        /// </summary>
        /// <returns>completiong of the method</returns>
        private bool ExtractMSI(string tempFolder)
        {
            bool isDone = false;
            if(strExeFilePath.Substring(strExeFilePath.Length-3,3).ToUpper() == "MSI")
            {
                isDone = true;
                return isDone;
            }
            var cmdString = string.Format(@"{0} /a /s /v""/qn TARGETDIR=\""{1}""""", strExeFilePath, tempFolder);
            ExecuteCommand(cmdString);

            isDone = true;
            return isDone;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Command"></param>
        private void ExecuteCommand(string Command)
        {
            Process myProcess = new Process();
            myProcess.StartInfo.FileName = "cmd.exe";
            myProcess.StartInfo.Arguments = "/c " + Command;
            myProcess.StartInfo.CreateNoWindow = true;
            myProcess.StartInfo.RedirectStandardOutput = false;
            myProcess.StartInfo.UseShellExecute = false;

            myProcess.Start();
            myProcess.WaitForExit();
            myProcess.Close();
        }


        private string GetMSIFile(string strFolderPath)
        {
            string[] filePaths = Directory.GetFiles(strFolderPath, "*.msi");
            if (filePaths.Length == 0)
                return null;
            return filePaths[0];
        }

        /// <summary>
        /// Gets the upgrade code (Product GUID) for the MSI file
        /// </summary>
        /// <param name="msiFileString"></param>
        /// <returns></returns>
        private bool GetMSIProperties(string msiFileString)
        {
            bool isDone = false;
            if (msiFileString == null)
            {
                Console.WriteLine("Unable to extract MSI");
                return isDone;
            }
            using (Database msidb = new Database(msiFileString, OpenDatabase.ReadOnly))
            {
                const string tableName = "Property";
                string query = string.Concat("SELECT * FROM `", tableName, "`");
                View propView = msidb.OpenExecuteView(query);
                Record columns  = propView.Fetch();
                
                while (columns != null)
                {
                    InstallerProps.Add(columns[1], columns[2]);
                    columns = propView.Fetch();
                }
                isDone = true;
            }

            return isDone;
        }
    }
}
