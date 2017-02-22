using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System.IO;
using System.Configuration;
using System.Runtime.InteropServices;
using Winsoft.Csv;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;
using System.Collections;

namespace ConsoleApplication1
{
    class Program
    {
        //NEW PARSING USING JETDB CONTINUATION

        public static void InitVar()
        {
            Logs.CreateLogFile();
            Config.dbHost = ConfigurationManager.AppSettings["Server"];
            Config.dbPort = ConfigurationManager.AppSettings["Port"];
            Config.dbName = ConfigurationManager.AppSettings["Database"];
            Config.dbUser = ConfigurationManager.AppSettings["UserName"];
            Config.dbPass = ConfigurationManager.AppSettings["Password"];
            Config.RawDir = ConfigurationManager.AppSettings["RawScanDir"];
            Config.ConvertedDir = ConfigurationManager.AppSettings["ConvertedScanDir"];
            Config.FileCount = Convert.ToInt32(ConfigurationManager.AppSettings["FileCount"]);
            Config.OutputDir = ConfigurationManager.AppSettings["OutputDir"];

            DirectoryInfo convertedDir = new DirectoryInfo(Config.ConvertedDir);
            DirectoryInfo outputDir = new DirectoryInfo(Config.OutputDir);
                foreach (FileInfo fi in convertedDir.GetFiles())
                {
                    Logs.Write(@"Deleting temp file: " + fi.Name);
                    fi.Delete();
                }


                foreach (FileInfo fi in outputDir.GetFiles())
                {
                    Logs.Write(@"Deleting temp file: " + fi.Name);
                    fi.Delete();
                }

        }

        static void Main(string[] args)
        {
            InitVar();
            DateTime startTime = DateTime.Now.ToLocalTime();
            string argMonth = string.Empty;
            string argYear = string.Empty;
            bool processRaw = true;

            foreach (string aw in args)
            {
                string[] temp;
                if (aw.Contains("-m"))
                {
                    temp = aw.Split('=');
                    argMonth = temp[1];
                }
                if (aw.Contains("-y"))
                {
                    temp = aw.Split('=');
                    argYear = temp[1];
                }
                if (aw.Contains("-p"))
                {
                    temp = aw.Split('=');
                    processRaw = Convert.ToBoolean(temp[1]);
                }
            }

            string formattedFile;
            string convertedFile;
            string fiDt;

            //DirectoryInfo rawReportsDir = new DirectoryInfo(System.Windows.Forms.Application.StartupPath + "\\RawReports");
            //DirectoryInfo convertedDir = new DirectoryInfo(System.Windows.Forms.Application.StartupPath + "\\Converted");
            DirectoryInfo rawReportsDir = new DirectoryInfo(Config.RawDir);
            DirectoryInfo convertedDir = new DirectoryInfo(Config.ConvertedDir);

            foreach (FileInfo fi in rawReportsDir.GetFiles())
            {
                string[] temp;
                Logs.Write("Formatting the file: " + fi.Name);
                fiDt = CsvToExcel(fi.FullName, convertedDir.FullName);
                Console.WriteLine("Format Complete!");
                temp = fiDt.Split(';');
                Logs.Write("Dumping: " + temp[0] + " to Database");
                parseFile(temp[0], true, temp[1]);
                Console.WriteLine("Dump Complete!");
            }

            //foreach (DictionaryEntry di in Db.GetRegion())
            //{
            //    Logs.Write("Generating Executive Report for: " + di.Value.ToString());
            //    Generate.ExecutiveReport(di.Value.ToString(), argMonth, argYear);
            //    Logs.Write("Done generating report for: " + di.Value.ToString());
            //}

            DateTime endTime = DateTime.Now.ToLocalTime();
            TimeSpan totalProcessTime = endTime.Subtract(startTime);
            Console.WriteLine(totalProcessTime.Minutes.ToString() + " Minutes");
            Console.WriteLine("Uwian Na!");
            Console.ReadKey(); 
        }

        public static string FormatTime(string date)
        {
            string dateTimeFormat = "yyyy-M-d HH:mm:ss";
            //Logs.Write("Formatting DateTime, format is " + dateTimeFormat + ".");
            DateTime dateTime = new DateTime();
            dateTime = Convert.ToDateTime(date);
            return dateTime.ToString(dateTimeFormat);
        }

        public static string CheckIfNull(OleDbDataReader reader, int index)
        {
            string result = string.Empty;

            if (reader.IsDBNull(index) == false)
            {
                string dataType = reader.GetDataTypeName(index);
                if (dataType == "DBTYPE_R8")
                {
                    result = Convert.ToString((reader.GetDouble(index)));
                }
                if (dataType == "DBTYPE_WVARCHAR")
                {
                    result = reader.GetString(index);
                }
            }
            string temp = Convert.ToString(result);
            temp = Sanitize(temp);
            return temp;
        }

        public static void parseFile(string fileName, bool isXlsFile, string dateTimeString)
        {
            FileInfo fi = new FileInfo(fileName);
            DateTime dt = Convert.ToDateTime(dateTimeString);

            string[] fName = fi.Name.Split('_');
            int region = Db.GetRegion(fName[0]);
            int buid = Db.GetBu(fName[1].Remove(fName[1].IndexOf('.')));

            int defaultOwnerID = 1;
            //int region = Config.RegionID;
            //int buid = Config.BuID;
            int defaultCi = 1;
            int assetID = 0;
            int vulnID = 0;

            //get first table for parsing
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fi.FullName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
            string wsName = xlWorksheet.Name;            
            xlWorkbook.Close();
            xlApp.Quit();

            OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fi.FullName + ";" + "Extended Properties='Excel 8.0;HDR=YES;IMEX=1;'");
            //OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fi.FullName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'");
            OleDbCommand cmd = new OleDbCommand(string.Format("Select * from [" + wsName + "$]"), con);                
            con.Open();
            OleDbDataReader reader = cmd.ExecuteReader();                    

            
            while (reader.Read())
            {
                Current.AssetID = 0;
                Current.VulnID = 0;
                Current.IPAddress = CheckIfNull(reader, 0);

                if(Current.IPAddress.Length > 15)
                {
                    break;
                }

                Current.Dns = CheckIfNull(reader, 1);
                Current.NetBios = CheckIfNull(reader, 2);
                Current.OperatingSystem = CheckIfNull(reader, 3);
                Current.FQDN = CheckIfNull(reader, 11);
                Current.QID = Convert.ToInt32(CheckIfNull(reader, 5));
                Current.Title = CheckIfNull(reader, 6);
                Current.VulnerabilityType = CheckIfNull(reader, 7);
                Current.Severity = Convert.ToInt32(CheckIfNull(reader, 8));
                Current.CVE = CheckIfNull(reader, 13);
                Current.VendorReference = CheckIfNull(reader, 14);
                Current.BugTraq_ID = CheckIfNull(reader, 15);
                Current.Threat = CheckIfNull(reader, 16);
                Current.Impact = CheckIfNull(reader, 17);
                Current.Solution = CheckIfNull(reader, 18);
                Current.Exploitability = CheckIfNull(reader, 19);
                Current.AssociatedMalware = CheckIfNull(reader, 20);
                Current.PCI = CheckIfNull(reader, 22);
                Current.Instance = CheckIfNull(reader, 23);
                Current.Category = CheckIfNull(reader, 24);
                Current.Port = CheckIfNull(reader, 9);
                Current.Protocol = CheckIfNull(reader, 10);


                if (Current.IPAddress == "10.3.3.41")
                {
                    Console.WriteLine();
                }

                Current.AssetID = Db.GetAssetID(defaultOwnerID, region, buid, defaultCi, Current.IPAddress, Current.Dns, Current.NetBios, Current.OperatingSystem);
                Current.VulnID = Db.GetVulnID(Current.Title, Current.QID, Current.VulnerabilityType, Current.Severity, Current.Threat, Current.Impact, Current.Solution, Current.Exploitability, Current.AssociatedMalware, Current.VendorReference, Current.CVE, Current.PCI, Current.Category, Current.Instance);
                //REMOVED BUGTRAQID
                //Current.VulnID = Db.GetVulnID(Current.Title, Current.QID, Current.VulnerabilityType, Current.Severity, Current.Threat, Current.Impact, Current.Solution, Current.Exploitability, Current.AssociatedMalware, Current.BugTraq_ID, Current.VendorReference, Current.CVE, Current.PCI, Current.Category, Current.Instance);
                if (Current.AssetID == 0)
                {
                    Db.InsertAssets(defaultOwnerID, region, buid, defaultCi, Current.IPAddress, Current.Dns, Current.NetBios, Current.OperatingSystem, Current.FQDN);
                }
                //refactored for reduced queries
                //if (Db.CheckIpIfExists(Current.IPAddress) == false)
                //{
                //    Db.InsertAssets(defaultOwnerID, region, buid, defaultCi, Current.IPAddress, Current.Dns, Current.NetBios, Current.OperatingSystem, Current.FQDN);
                //}
                if (Current.VulnID == 0)
                {
                    Db.InsertVuln(Current.Title, Current.QID, Current.VulnerabilityType, Current.Severity, Current.Threat, Current.Impact, Current.Solution, Current.Exploitability, Current.AssociatedMalware, Current.BugTraq_ID, Current.VendorReference, Current.CVE, Current.PCI, Current.Category, Current.Instance);
                }
                //refactored for reduced queries
                //if (Db.CheckIfVulnExists(Current.Title, Current.QID, Current.VulnerabilityType, Current.Severity, Current.Threat, Current.Impact, Current.Solution, Current.Exploitability, Current.AssociatedMalware, Current.BugTraq_ID, Current.VendorReference, Current.CVE, Current.PCI, Current.Category, Current.Instance) == false)
                //{
                //    Db.InsertVuln(Current.Title, Current.QID, Current.VulnerabilityType, Current.Severity, Current.Threat, Current.Impact, Current.Solution, Current.Exploitability, Current.AssociatedMalware, Current.BugTraq_ID, Current.VendorReference, Current.CVE, Current.PCI, Current.Category, Current.Instance);
                //}
                if(Current.AssetID == 0)
                {
                    Current.AssetID = Db.GetAssetID(defaultOwnerID, region, buid, defaultCi, Current.IPAddress, Current.Dns, Current.NetBios, Current.OperatingSystem);
                }
                if (Current.VulnID == 0)
                {
                    //REMOVED BUGTRAQID
                    //Current.VulnID = Db.GetVulnID(Current.Title, Current.QID, Current.VulnerabilityType, Current.Severity, Current.Threat, Current.Impact, Current.Solution, Current.Exploitability, Current.AssociatedMalware, Current.BugTraq_ID, Current.VendorReference, Current.CVE, Current.PCI, Current.Category, Current.Instance);
                    Current.VulnID = Db.GetVulnID(Current.Title, Current.QID, Current.VulnerabilityType, Current.Severity, Current.Threat, Current.Impact, Current.Solution, Current.Exploitability, Current.AssociatedMalware, Current.VendorReference, Current.CVE, Current.PCI, Current.Category, Current.Instance);
                }
                //

                Current.FID = Db.GetFindingsID(Current.AssetID, Current.VulnID, Current.Port, Current.Protocol,  Current.Results, Current.Status, dt.Month, dt.Year);

                if(Current.FID == 0)
                {
                    Db.InsertFindings(Current.AssetID, Current.VulnID, Current.Port, Current.Protocol, Current.Results, Current.Status, dt.Month, dt.Year);
                }
                
                //Current.AssetID, Current.VulnID, Current.Port, Current.Results, Current.Solution, Current.Status
            }                
            con.Close();            
        }

        public static string Sanitize(string line)
        {
            string result = line;

            if (result.Contains("'"))
            {
                result = result.Replace("'", string.Empty);
            }
            if (result.Contains("\""))
            {
                result = result.Replace("\"", string.Empty);
            }
            
            return result;
        }

        public static string CsvToExcel(string fileName, string dir)
        {
            string result;
            string dateTimeString;
            FileInfo fi = new FileInfo(fileName);
            string xlFileName = fi.Name.Remove(fi.Name.IndexOf(".")) + ".xlsx";
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fi.FullName, 0, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            xlWorkbook.SaveAs(dir + xlFileName, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //removes headers
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
            dateTimeString = xlWorksheet.Cells[6, 1].Text.ToString();
            dateTimeString = dateTimeString.Substring(0, dateTimeString.IndexOf("at") - 1);

            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.get_Range("A1", "A7");
            Microsoft.Office.Interop.Excel.Range row = xlRange.EntireRow;
            row.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);





            //remove footer if contains hosts not scanned, host not alive
            Microsoft.Office.Interop.Excel.Range xlRange1 = xlWorksheet.UsedRange;

            //cell pointer
            int rowCount = xlRange1.Rows.Count;
            int colCount = xlRange1.Columns.Count;

            Microsoft.Office.Interop.Excel.Range xlRange2 = xlWorksheet.get_Range("A" + rowCount, "A" + rowCount);
            Microsoft.Office.Interop.Excel.Range lastRow  = xlRange2.EntireRow;
            if(xlWorksheet.Cells[rowCount,5].Text.ToString().Contains("hosts not scanned, host not alive"))
            {
                lastRow.Delete();
                Logs.Write(fi.Name);
            }
            

            xlApp.DisplayAlerts = false;
            xlWorkbook.Save();
            xlWorkbook.Close();
            xlApp.Quit();
            return result = dir + "\\" + xlFileName + ";" + dateTimeString;
            //return result;
        }

        public static void ExcelToCSV(string fileName)
        {
            FileInfo fi = new FileInfo(fileName);
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fi.FullName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            xlWorkbook.SaveAs(fi.DirectoryName + "\\TheShit.xlsx", XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            xlWorkbook.Close();
            xlApp.Quit();
        }

        public static string ExcelToCSV(string fileName, bool removeHeaders)
        {
            string result;
            FileInfo fi = new FileInfo(fileName);
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fi.FullName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.get_Range("A1", "A7");
            Microsoft.Office.Interop.Excel.Range row = xlRange.EntireRow;
            row.Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);

            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(fi.DirectoryName + "\\" + fi.Name.Remove(fi.Name.IndexOf(".")) + ".csv", XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            xlWorkbook.Close();
            xlApp.Quit();
            return result = fi.FullName;
        }
        
    }
}
