using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System.Collections;


//business units
//active hosts
//top 10 sev 4&5 vuln


namespace ConsoleApplication1
{
    class Generate
    {
        public static void ExecutiveReport(string region, string month, string year)
        {
            #region Variable Declaration

            string dummyDate = month + " 1, " + year + " 12:00:00AM";
            DateTime repDate = Convert.ToDateTime(dummyDate);
            month = repDate.ToString("MMMM");
            string sev1 = Db.GetSeverityCount(1, repDate.Month, repDate.Year, region).ToString();
            string sev2 = Db.GetSeverityCount(2, repDate.Month, repDate.Year, region).ToString();
            string sev3 = Db.GetSeverityCount(3, repDate.Month, repDate.Year, region).ToString();
            string sev4 = Db.GetSeverityCount(4, repDate.Month, repDate.Year, region).ToString();
            string sev5 = Db.GetSeverityCount(5, repDate.Month, repDate.Year, region).ToString();
            string template = @"C:\Users\IEUser\Downloads\Executive Report_Template.docx"; //Executive Report Template
            FileInfo pieChart = new FileInfo(@"C:\Users\IEUser\Downloads\pieChart.xlsx");
            string regionPH = "Vregion"; //Region placeholder
            string monthPH = "Vmonth"; //Month placeholder
            string yearPH = "Vyear"; //Year placeholder
            string bunitPH = "Vbu"; //Business unit placeholder
            string hostPH = "Vhost"; //Active hosts placeholder
            string critPH = "Vcrit"; //Total critical (sev4&5) vuln count place holder
            string vulnPH = "Vvuln"; //Total vuln count (sev1-5) placeholder
            string sev5PH = "Sev5"; //Count of Sev 5 vuln
            string sev4PH = "Sev4"; //Count of Sev 4 vuln
            string sev3PH = "Sev3"; //Count of Sev 3 vuln
            string sev2PH = "Sev2"; //Count of Sev 2 vuln
            string sev1PH = "Sev1"; //Count of Sev 1 vuln
            string Top01 = "Top01"; //Top Vulnerabilities 1-10; Sev 4&5 only
            string TSev01 = "TSev01"; //Severity of Top 1-10 vulnerability
            string Tcount01 = "Tcount01"; //Count of each top 1-10 vulnerability
            int totalVulnCount = Convert.ToInt32(sev1) + Convert.ToInt32(sev2) + Convert.ToInt32(sev3) + Convert.ToInt32(sev4) + Convert.ToInt32(sev5);
            int critVuln = Convert.ToInt32(sev4) + Convert.ToInt32(sev5);
            int activeHosts = Db.GetActiveHost(repDate.Month, repDate.Year, region);

            #endregion

            //Excel Workbook Modification
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(pieChart.FullName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            #region Pie Chart

            //Db.GetSeverityCount(int severity, int month, int year, string region)
            xlWorksheet.Cells[2, 2] = sev1;
            xlWorksheet.Cells[3, 2] = sev2;
            xlWorksheet.Cells[4, 2] = sev3;
            xlWorksheet.Cells[5, 2] = sev4;
            xlWorksheet.Cells[6, 2] = sev5;

            #endregion

            #region Column Chart

            string[] Assets;
            int row = 19;

            foreach (DictionaryEntry di in Db.GetTopVulnAssests(region, repDate.Month, repDate.Year))
            {
                Assets = di.Value.ToString().Split('|');
                bool flag = true;
                foreach (string temp in Assets)
                {
                    if (flag)
                    {
                        xlWorksheet.Cells[row, 1] = di.Key.ToString();
                        flag = false;
                    }
                    else
                    {
                        xlWorksheet.Cells[row, 1] = "";
                    }
                    string[] sep = temp.Split(':');
                    xlWorksheet.Cells[row, 2] = sep[0];
                    xlWorksheet.Cells[row, 3] = sep[1];
                    row++;
                }
            }

            #endregion

            xlWorkbook.Save();

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            //Microsoft.Office.Interop.Word.Document execReport = wordApp.Documents.Open(template);
            Microsoft.Office.Interop.Word.Document execReport = wordApp.Documents.Open(template, false, false, false, Format: Microsoft.Office.Interop.Word.WdOpenFormat.wdOpenFormatXMLDocument);
            wordApp.Visible = false;

            //Replace placeholder
            Microsoft.Office.Interop.Word.Find findObject = wordApp.Selection.Find;
            object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;

            #region Replace Placeholders

            findObject.Execute(regionPH, true, true, false, false, false, true, Type.Missing, true, region, WdReplace.wdReplaceAll, false, false, false, false);
            findObject.Execute(monthPH, true, true, false, false, false, true, Type.Missing, true, month, WdReplace.wdReplaceAll, false, false, false, false);
            findObject.Execute(yearPH, true, true, false, false, false, true, Type.Missing, true, year, WdReplace.wdReplaceAll, false, false, false, false);
            findObject.Execute(sev1PH, true, true, false, false, false, true, Type.Missing, true, sev1, WdReplace.wdReplaceAll, false, false, false, false);
            findObject.Execute(sev2PH, true, true, false, false, false, true, Type.Missing, true, sev2, WdReplace.wdReplaceAll, false, false, false, false);
            findObject.Execute(sev3PH, true, true, false, false, false, true, Type.Missing, true, sev3, WdReplace.wdReplaceAll, false, false, false, false);
            findObject.Execute(sev4PH, true, true, false, false, false, true, Type.Missing, true, sev4, WdReplace.wdReplaceAll, false, false, false, false);
            findObject.Execute(sev5PH, true, true, false, false, false, true, Type.Missing, true, sev5, WdReplace.wdReplaceAll, false, false, false, false);
            findObject.Execute(vulnPH, true, true, false, false, false, true, Type.Missing, true, totalVulnCount.ToString(), WdReplace.wdReplaceAll, false, false, false, false);
            findObject.Execute(critPH, true, true, false, false, false, true, Type.Missing, true, critVuln.ToString(), WdReplace.wdReplaceAll, false, false, false, false);
            findObject.Execute(hostPH, true, true, false, false, false, true, Type.Missing, true, activeHosts, WdReplace.wdReplaceAll, false, false, false, false);

            #endregion

            #region Generete Charts

            //select pie chart from xlworksheet
            Microsoft.Office.Interop.Excel.ChartObject pchart = xlWorksheet.ChartObjects("PieChart"); //PieChart = name of chart; returns chart objects
            Microsoft.Office.Interop.Excel.Chart piechart = pchart.Chart; //return actual chart
            piechart.ChartArea.Select(); // select and copy whole chart area
            piechart.ChartArea.Copy();

            //position cursor before chart insert
            execReport.Application.Selection.Find.Execute("<PieChart>");

            //paste copied chart to word
            execReport.Application.Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdFormatOriginalFormatting);

            //select pie chart from xlworksheet
            Microsoft.Office.Interop.Excel.ChartObject cchart = xlWorksheet.ChartObjects("ColumnChart"); //ColumnChart = name of chart; returns chart objects
            Microsoft.Office.Interop.Excel.Chart colchart = cchart.Chart; //return actual chart
            colchart.ChartArea.Select(); // select and copy whole chart area
            colchart.ChartArea.Copy();

            //position cursor before chart insert
            execReport.Application.Selection.Find.Execute("<ColumnChart>");

            //paste copied chart to word
            execReport.Application.Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdFormatOriginalFormatting);

            #endregion

            //close workbook
            xlWorkbook.Close(false, pieChart.FullName);
            xlApp.Quit();

            //Save modified word doc
            string fileName = Config.OutputDir + "\\Executive Report_" + region + "_" + month + year + ".docx";
            execReport.SaveAs(fileName);

            //save as PDF
            string pdfname = fileName.Replace(".docx", ".pdf");
            execReport.ExportAsFixedFormat(pdfname, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, false);

            execReport.Close();
            wordApp.Quit();
        }
    }
}
