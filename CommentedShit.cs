using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ConsoleApplication1
{
    class CommentedShit
    {
        //public static void parseFile(string fileName)
        //{
        //    FileInfo fi = new FileInfo(fileName);
        //    int defaultOwnerID = 1;
        //    int region = Config.RegionID;
        //    int buid = Config.BuID;
        //    int defaultCi = 1;
        //    int assetID = 0;
        //    int vulnID = 0;

        //    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        //    //Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.OpenText(fi.FullName, Type.Missing,2, XlTextParsingType.xlDelimited, XlTextQualifier.xlTextQualifierDoubleQuote, false, false, false, true, false, false, false, XlColumnDataType.xlTextFormat, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        //    //Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.OpenText(fi.FullName, DataType:XlTextParsingType.xlDelimited, Comma: true);
        //    Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fi.FullName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        //    Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
        //    Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;


        //    //handling null value for region and buid
        //    if (Config.RegionID == 0 || Config.BuID == 0)
        //    {
        //        Console.WriteLine("Region or Business Unit does not exists!!");
        //        Console.ReadKey();
        //        Environment.Exit(69);
        //    }

        //    //-1 para d kunin ung last row
        //    int rowCount = xlRange.Rows.Count - 1;
        //    int colCount = xlRange.Columns.Count;

        //    //start sa relevant info
        //    for (int row = 1; row <= rowCount; row++)
        //    {
        //        Current.IPAddress = xlWorksheet.Cells[row, 1].Text.ToString();
        //        Current.Dns = xlWorksheet.Cells[row, 2].Text.ToString();
        //        Current.NetBios = xlWorksheet.Cells[row, 3].Text.ToString();
        //        Current.OperatingSystem = xlWorksheet.Cells[row, 4].Text.ToString();
        //        Current.FQDN = xlWorksheet.Cells[row, 12].Text.ToString();


        //        //if (Db.CheckIpIfExists(Current.IPAddress) == false)
        //        //{
        //        Db.InsertAssets(defaultOwnerID, region, buid, defaultCi, Current.IPAddress, Current.Dns, Current.NetBios, Current.OperatingSystem, Current.FQDN, defaultCi);
        //        //}
        //        //assetID = Db.GetAssetID(defaultOwnerID, region, buid, defaultCi, Current.IPAddress, Current.Dns, Current.NetBios, Current.OperatingSystem, Current.FQDN, defaultCi);


        //        Current.QID = Convert.ToInt32(xlWorksheet.Cells[row, 6].Text.ToString());
        //        Current.Title = Sanitize(xlWorksheet.Cells[row, 7].Text.ToString());
        //        Current.VulnerabilityType = Sanitize(xlWorksheet.Cells[row, 8].Text.ToString());
        //        Current.Severity = Convert.ToInt32(xlWorksheet.Cells[row, 9].Text.ToString());
        //        Current.CVE = Sanitize(xlWorksheet.Cells[row, 14].Text.ToString());
        //        Current.VendorReference = Sanitize(xlWorksheet.Cells[row, 15].Text.ToString());
        //        Current.BugTraq_ID = Sanitize(xlWorksheet.Cells[row, 16].Text.ToString());
        //        Current.Threat = Sanitize(xlWorksheet.Cells[row, 17].Text.ToString());
        //        Current.Impact = Sanitize(xlWorksheet.Cells[row, 18].Text.ToString());
        //        Current.Solution = Sanitize(xlWorksheet.Cells[row, 19].Text.ToString());
        //        Current.Exploitability = Sanitize(xlWorksheet.Cells[row, 20].Text.ToString());
        //        Current.AssociatedMalware = Sanitize(xlWorksheet.Cells[row, 21].Text.ToString());
        //        Current.PCI = Sanitize(xlWorksheet.Cells[row, 23].Text.ToString());
        //        Current.Instance = Sanitize(xlWorksheet.Cells[row, 24].Text.ToString());
        //        Current.Category = Sanitize(xlWorksheet.Cells[row, 25].Text.ToString());



        //        //if (Db.CheckIfVulnExists(Current.Title, Current.QID, Current.VulnerabilityType, Current.Severity, Current.Threat, Current.Impact, Current.Solution, Current.Exploitability, Current.AssociatedMalware, Current.BugTraq_ID, Current.VendorReference, Current.CVE, Current.PCI, Current.Category, Current.Instance) == false)
        //        //{
        //        Db.InsertVuln(Current.Title, Current.QID, Current.VulnerabilityType, Current.Severity, Current.Threat, Current.Impact, Current.Solution, Current.Exploitability, Current.AssociatedMalware, Current.BugTraq_ID, Current.VendorReference, Current.CVE, Current.PCI, Current.Category, Current.Instance);
        //        //}

        //        //Console.ReadKey();								  


        //        //Console.WriteLine(Db.CheckIpIfExists(Current.IPAddress));
        //        //Display Assets

        //        //No IP Status field
        //        //Console.WriteLine("IP Status: " + xlWorksheet.Cells[row, 5].Text.ToString());

        //        //Console.WriteLine("QID: " + xlWorksheet.Cells[row, 6].Text.ToString());
        //        //Console.WriteLine("Title: " + xlWorksheet.Cells[row, 7].Text.ToString());
        //        //Console.WriteLine("Type: " + xlWorksheet.Cells[row, 8].Text.ToString());
        //        //Console.WriteLine("Severity: " + xlWorksheet.Cells[row, 9].Text.ToString());
        //        //Console.WriteLine("Port: " + xlWorksheet.Cells[row, 10].Text.ToString());

        //        //Console.WriteLine("Protocol: " + xlWorksheet.Cells[row, 11].Text.ToString());

        //        //Console.WriteLine("SSL: " + xlWorksheet.Cells[row, 13].Text.ToString());
        //        //Console.WriteLine("CVE ID: " + xlWorksheet.Cells[row, 14].Text.ToString());
        //        //Console.WriteLine("Vendor Reference: " + xlWorksheet.Cells[row, 15].Text.ToString());
        //        //Console.WriteLine("Bugtraq ID: " + xlWorksheet.Cells[row, 16].Text.ToString());
        //        //Console.WriteLine("Threat: " + xlWorksheet.Cells[row, 17].Text.ToString());
        //        //Console.WriteLine("Impact: " + xlWorksheet.Cells[row, 18].Text.ToString());
        //        //Console.WriteLine("Solution: " + xlWorksheet.Cells[row, 19].Text.ToString());
        //        //Console.WriteLine("Exploitability: " + xlWorksheet.Cells[row, 20].Text.ToString());

        //        //Console.WriteLine("Associated Malware: " + xlWorksheet.Cells[row, 21].Text.ToString());
        //        //Console.WriteLine("Results: " + xlWorksheet.Cells[row, 22].Text.ToString());
        //        //Console.WriteLine("PCI Vuln: " + xlWorksheet.Cells[row, 23].Text.ToString());
        //        //Console.WriteLine("Instance: " + xlWorksheet.Cells[row, 24].Text.ToString());
        //        //Console.WriteLine("Category: " + xlWorksheet.Cells[row, 25].Text.ToString());


        //        //Db.InsertAssets(defaultOwnerID, region, buid, defaultCi, Current.IPAddress, Current.Dns, Current.NetBios, Current.OperatingSystem, Current.FQDN,defaultCi);
        //        //Console.WriteLine("********************************END OF SHIT **************************************");
        //        //Console.ReadKey();
        //    }
        //    xlWorkbook.Close();
        //    xlApp.Quit();
        //}

        //public static void parseFile(string fileName, bool isCsv, string sheetName)
        //{
        //    FileInfo file = new FileInfo(fileName);

        //    //get first table for parsing
        //    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file.FullName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        //    Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
        //    string wsName = xlWorksheet.Name;
        //    xlWorkbook.Close();
        //    xlApp.Quit();

        //    //using (OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"" + file.DirectoryName + "\";Extended Properties='text;HDR=Yes;FMT=Delimited(,)';"))
        //    //ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";


        //    using (OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file.FullName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"))
        //    {
        //        //using (OleDbCommand cmd = new OleDbCommand(string.Format("SELECT * FROM [{0}]", "Australia_Sydney"), con))
        //        using (OleDbCommand cmd = new OleDbCommand(string.Format("Select * from [" + wsName + "$]"), con))
        //        {
        //            con.Open();

        //            // Using a DataReader to process the data
        //            using (OleDbDataReader reader = cmd.ExecuteReader())
        //            {
        //                while (reader.Read())
        //                {
        //                    Console.WriteLine(reader.GetString(0));
        //                    Console.ReadKey();
        //                }
        //            }

        //             //Using a DataTable to process the data
        //            //using (OleDbDataAdapter adp = new OleDbDataAdapter(cmd))
        //            //{
        //            //    System.Data.DataTable tbl = new System.Data.DataTable("MyTable");
        //            //    adp.Fill(tbl);

        //            //    foreach (DataRow row in tbl.Rows)
        //            //    {
        //            //        // Process the current row...
        //            //    }
        //            //}
        //        }
        //        con.Close();
        //    }
        //}


        //public static string CheckIfNull(OleDbDataReader reader, int index)
        //{
        //    var result = string.Empty;

        //    if(reader.IsDBNull(index) == false)
        //    {
        //        result = reader.GetString(index);
        //    }

        //    result = Sanitize(result);
        //    return result;
        //}

        //public static string CheckIfNull(OleDbDataReader reader, int index, bool isFloat)
        //{
        //    double result = 0;

        //    if (reader.IsDBNull(index) == false)
        //    {
        //        result = reader.GetDouble(index);
        //    }
        //    string temp = Convert.ToString(result);
        //    temp = Sanitize(temp);
        //    return temp;
        //}



        //foreach (FileInfo fi in convertedDir.GetFiles())
        //{
        //    Console.WriteLine("Dumping: " + fi.Name + " to Database");
        //    parseFile(fi.FullName, true, ht.);
        //    Console.WriteLine("Dump Complete!");
        //}

        //Console.WriteLine("Formatting the file...");
        //formattedFile = CsvToExcel(args[0]);
        //Console.WriteLine("Format Complete!");

        //CsvToExcel(@"C:\Users\IEUser\Documents\GitHub\AyVeeEm\ConsoleApplication1\bin\Debug\RawReports\South Africa_Dar Es Salaam.csv", convertedDir.FullName);
        //parseFile(@"C:\Users\IEUser\Documents\GitHub\AyVeeEm\ConsoleApplication1\bin\Debug\Converted\South Africa_Dar Es Salaam.xlsx", true, "12/10/2016");
        //Console.WriteLine("Converting the File...");
        //convertedFile = ExcelToCSV(formattedFile, true);
        //Console.WriteLine("Finished Converting");
        //CsvToExcel(@"C:\Users\IEUser\Documents\GitHub\AyVeeEm\ConsoleApplication1\bin\Debug\RawReports\South Africa_Dar Es Salaam.csv", convertedDir.FullName);
        //parseFile(@"C:\Users\IEUser\Downloads\TestData\EMEA_London WGC.xlsx", true);

        //Console.WriteLine(startTime);
    }
}
