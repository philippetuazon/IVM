using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;
using System.IO;
using System.Collections;

namespace ConsoleApplication1
{
    class Db
    {
        static string connectionString = "Server=" + Config.dbHost + ";Port=" + Config.dbPort + ";Database=" + Config.dbName + ";Uid=" + Config.dbUser + ";Password=" + Config.dbPass;

        //function for testing function
        public static void testConnection()
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            try
            {
                conn.Open();
                Console.WriteLine("Connection to DB Success! :p");
                conn.Close();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex);
            }
        }
        //returns region id in int of the passed region in string
        public static int GetRegion(string region)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            int result = 0;
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            command.CommandText = "SELECT * FROM Region WHERE Region = \"" + region + "\"";
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                result = Convert.ToInt32(dr.GetString(0));
                Console.WriteLine("Region ID for " + region + " is: " + result);
            }
            conn.Close();

            return result;
        }
        //returns businessunit id in int of the passed businessunit in string
        public static int GetBu(string bu)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            int result = 0;
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            command.CommandText = "SELECT * FROM Business_Unit WHERE BUnit = \"" + bu + "\"";
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                result = Convert.ToInt32(dr.GetString(0));
                Console.WriteLine("Business ID for " + bu + " is: " + result);
            }
            conn.Close();

            return result;
        }
        //returns the asset id in int of the passed ipaddress in string
        public static bool CheckIpIfExists(string IP)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            bool result = false;
            int ipId = 0;
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            //Console.WriteLine("Checking for duplicate IP: " + IP);
            command.CommandText = "SELECT * FROM assets WHERE IP_Address = \"" + IP + "\"";
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                ipId = Convert.ToInt32(dr.GetString(0));
            }
            conn.Close();
            if(ipId != 0)
            {
                result = true;
            }
            return result;
        }
        //for populating assests table
        public static void InsertAssets(int ownerID, int regionID, int buID, int defaultCi, string IP, string dns, string netBios, string operatingSystem, string fqdn)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            int result = 0;
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            //Console.WriteLine("Inserting asset values for: " + IP);

            command.CommandText = "INSERT INTO ASSETS (Owner_ID, Region_ID, BUnit_ID, IP_Address, DNS, NetBios, OperatingSystem, FQDN, Ci) VALUES ('" + ownerID + "','" + regionID + "','" + buID + "','" + IP + "','" + dns + "','" + netBios + "','" + operatingSystem + "','" + fqdn + "','" + defaultCi + "')";
            command.ExecuteNonQuery();
            conn.Close();
        }
        //returns the asset ID
        public static int GetAssetID(int ownerID, int regionID, int buID, int defaultCi, string IP, string dns, string netBios, string operatingSystem)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            int assetID = 0;


            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            command.CommandText = "SELECT Asset_ID FROM assets WHERE Owner_ID = \"" + ownerID + "\" AND Region_ID = \"" + regionID + "\" AND BUnit_ID = \"" + buID + "\" AND IP_Address = \"" + IP + "\" AND DNS = \"" + dns + "\" AND NetBios = \"" + netBios + "\" AND OperatingSystem = \"" + operatingSystem + "\" AND Ci = \"" + defaultCi + "\"";
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                assetID = Convert.ToInt32(dr.GetString(0));
            }
            conn.Close();
            return assetID;
        }

        public static int GetVulnID(string title, int qid, string vtype, int severity, string threat, string impact, string solution, string exploitability, string assocMalware, string vendorRef, string cve, string pci, string category, string instance)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            int vulnID = 0;
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            command.CommandText = "SELECT Vuln_ID FROM Vulnerability WHERE Title = \"" + title + "\" AND QID = \"" + qid + "\" AND VType = \"" + vtype + "\" AND Severity = \"" + severity + "\" AND Threat = \"" + threat + "\" AND Impact = \"" + impact + "\" AND Solution = \"" + solution + "\" AND Exploitability = \"" + exploitability + "\" AND Assoc_Malware = \"" + assocMalware + "\" AND VendorRef = \"" + vendorRef + "\" AND CVE = \"" + cve + "\" AND PCI = \"" + pci + "\" AND Category = \"" + category + "\" AND Instance = \"" + instance + "\"";
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                vulnID = Convert.ToInt32(dr.GetString(0));
            }
            conn.Close();
            return vulnID;
        }

        public static string TestSelectQuery(string val)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            //string line = string.Empty;
            string line = string.Empty;
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            command.CommandText = "SELECT UNHEX(Exploitability) FROM Vulnerability WHERE Vuln_ID = \"" + val + "\"";
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                Console.WriteLine(line = dr.GetString(0));
            }
            conn.Close();

            return line;
            //if query returned value that is not 0 ip address already exists
        }

        public static void TestInsertQuery(string val, string val1)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            int result = 0;
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            val = Program.Sanitize(val);
            Console.WriteLine("Inserting Value: " + val);
            command.CommandText = "INSERT INTO Vulnerability (Title, Threat) VALUES (HEX('" + val + "'),HEX('" + val1 + "'))";
            command.ExecuteNonQuery();
            conn.Close();
        }
        //WALA PALNG SILBE
        public static void UpdateAsset(int assetID, int regionid, int bu)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            int result = 0;
            //
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            command.CommandText = "SELECT * FROM testTable WHERE Asset_ID = " + assetID;
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                result = Convert.ToInt32(dr.GetString(0));
                Console.WriteLine(result);
            }
        }
        public static bool CheckIfVulnExists(string title, int qid, string vtype, int severity, string threat, string impact, string solution, string exploitability, string assocMalware, string vendorRef, string cve, string pci, string category, string instance)
        //public static bool CheckIfVulnExists(string title, int qid, string vtype, int severity, string threat, string impact, string solution, string exploitability, string assocMalware, string bugTraqID, string vendorRef, string cve, string pci, string category, string instance)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            bool result = false;
            int hit = 0;
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            command.CommandText = "SELECT * FROM Vulnerability WHERE Title = \"" + title + "\" AND QID = \"" + qid + "\" AND VType = \"" + vtype + "\" AND Severity = \"" + severity + "\" AND Threat = \"" + threat + "\" AND Impact = \"" + impact + "\" AND Solution = \"" + solution + "\" AND Exploitability = \"" + exploitability + "\" AND Assoc_Malware = \"" + assocMalware + "\" AND VendorRef = \"" + vendorRef + "\" AND CVE = \"" + cve + "\" AND PCI = \"" + pci + "\" AND Category = \"" + category + "\" AND Instance = \"" + instance + "\"";
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                hit = Convert.ToInt32(dr.GetString(0));
            }
            conn.Close();
            //if query returned value that is not 0 ip address already exists

            if (hit != 0)
            {
                result = true;
            }
            return result;
        }

        public static void InsertVuln(string title, int qid, string vtype, int severity, string threat, string impact, string solution, string exploitability, string assocMalware, string bugTraqID, string vendorRef, string cve, string pci, string category, string instance)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            int result = 0;
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            //Console.WriteLine("Adding entry for: " + title);
            //command.CommandText = "INSERT INTO Vulnerability (Title, QID, Type, Severity, Threat, Impact, Solution, Exploitability, Assoc_Malware, BugTraqID, VendorRef, CVE, PCI, Category, Instance) VALUES ('" + title + "','" + qid + "','" + type + "','" + severity + "',HEX('" + threat + "'),HEX('" + impact + "'),HEX('" + solution + "'),HEX('" + exploitability + "'),'" + assocMalware + "','" + bugTraqID + "','" + vendorRef + "','" + cve + "','" + pci + "','" + category + "','" + instance + ")";
            command.CommandText = "INSERT INTO Vulnerability (Title, QID, VType, Severity, Threat, Impact, Solution, Exploitability, Assoc_Malware, BugTraqID, VendorRef, CVE, PCI, Category, Instance) VALUES ('" + title + "','" + qid + "','" + vtype + "', '" + severity + "', '" + threat + "', '" + impact + "', '" + solution + "', '" + exploitability + "','" + assocMalware + "','" + bugTraqID + "','" + vendorRef + "','" + cve + "','" + pci + "','" + category + "','" + instance + "')";
            command.ExecuteNonQuery();
            conn.Close();
        }

        public static void InsertFindings(int assetid, int vulnid, string port, string protocol, string results, string currentstatus, int month, int year)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            int result = 0;
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            command.CommandText = "INSERT INTO Findings (Asset_ID, Vuln_ID, Port, Protocol, Results, Status, Month, Year) VALUES ('" + assetid + "', '" + vulnid + "', '" + port + "', '" + protocol + "', '" + results + "', '" + currentstatus + "', '" + month + "', '" + year + "')";
            command.ExecuteNonQuery();
            conn.Close();
        }

        public static int GetFindingsID(int assetid, int vulnid, string port, string protocol, string results, string currentstatus, int month, int year)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            int vulnID = 0;
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            command.CommandText = "SELECT F_ID FROM Findings WHERE Asset_ID = \"" + assetid + "\" AND Vuln_ID = \"" + vulnid + "\" AND Port = \"" + port + "\" AND Protocol = \"" + protocol + "\" AND Results = \"" + results + "\"AND Status = \"" + currentstatus + "\"AND Month = \"" + month + "\"AND Year = \"" + year + "\"";
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                vulnID = Convert.ToInt32(dr.GetString(0));
            }
            conn.Close();
            return vulnID;
        }


        public static int GetSeverityCount(int severity, int month, int year, string region)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            int result = 0;
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
            command.CommandText = "SELECT COUNT(f.F_ID) " +
                                  "FROM Findings f, Vulnerability v, Assets a, Region r " +
                                  "WHERE f.Vuln_ID = v.Vuln_ID " +
                                  "AND f.Asset_ID = a.Asset_ID " +
                                  "AND a.Region_ID = r.Region_ID " +
                                  "AND v.Severity = " + severity + " " +
                                  "AND f.Year = " + year + " " +
                                  "AND f.Month = " + month + " " +
                                  "AND v.VType != \"Ig\" " +
                                  "AND r.Region = \"" + region + "\"";
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                result = Convert.ToInt32(dr.GetString(0));
            }

            conn.Close();
            return result;  
        }

        public static Hashtable GetRegion()
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            //string line = string.Empty;
            Hashtable ht = new Hashtable();
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

            command.CommandText = "SELECT Region_ID, Region FROM Region";
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                ht.Add(dr.GetString(0), dr.GetString(1));                
            }
            conn.Close();

            return ht;
        }        

        public static int GetActiveHost(int month, int year, string region)
        {
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            int result = 0;
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
            command.CommandText = "SELECT COUNT(DISTINCT a.Asset_ID) " +
                                  "FROM Findings a, Assets b, Region r " +
                                  "WHERE r.Region_ID = b.Region_ID " +
                                  "AND a.Asset_ID = b.Asset_ID " +
                                  "AND a.Month = " + month + " " +
                                  "AND a.Year = " + year + " " +
                                  "AND r.Region = \"" + region + "\"";

            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                result = Convert.ToInt32(dr.GetString(0));
            }

            conn.Close();
            return result;
        }

        public static Hashtable GetTopVulnAssests(string region, int month, int year)
        {
            Hashtable topAssets = new Hashtable();
            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
            command.CommandText = "SELECT a.OperatingSystem, a.IP_Address,  count(f.Vuln_ID) " +
                                  "FROM ivm_nrf.Findings f, ivm_nrf.Vulnerability v, ivm_nrf.Assets a, ivm_nrf.Region r " +
                                  "WHERE f.Asset_ID = a.Asset_ID " +
                                  "AND f.Vuln_ID = v.Vuln_ID " +
                                  "AND a.Region_ID = r.Region_ID " +
                                  "AND v.Severity in (4,5) " +
                                  "AND v.VType != \"Ig\" " +
                                  "AND r.Region = \"" + region + "\" " +
                                  "AND f.Year = " + year + " " +
                                  "AND f.Month = " + month + " " +
                                  "group by a.IP_Address " +
                                  "order by count(f.Vuln_ID) desc " +
                                  "limit 10";
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                if (topAssets.ContainsKey(dr.GetString(0)))
                {
                    topAssets[dr.GetString(0)] += "|" + dr.GetString(1) + ":" + dr.GetString(2);
                }
                else
                {
                    topAssets.Add(dr.GetString(0), dr.GetString(1) + ":" + dr.GetString(2));
                }
            }

            conn.Close();

            return topAssets;
        }

        //UNTESTED

        public static Hashtable GetMonthlyData(string month, string year, string region)
        {
            Hashtable ht = new Hashtable();

            MySqlConnection conn = new MySqlConnection(connectionString);
            MySqlCommand command = conn.CreateCommand();
            try
            {
                conn.Open();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
            command.CommandText = "";
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
            {
                //result = Convert.ToInt32(dr.GetString(0));
            }

            conn.Close();
            return ht;  
        }
    }
}
