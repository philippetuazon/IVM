using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ConsoleApplication1
{
    class Current
    {
        #region ID's

        static int assetid;
        static int vulnid;
        static int fid;


        public static int AssetID
        {
            get
            {
                return assetid;
            }
            set
            {
                assetid = value;
            }
        }

        public static int VulnID
        {
            get
            {
                return vulnid;
            }
            set
            {
                vulnid = value;
            }
        }

        public static int FID
        {
            get
            {
                return fid;
            }
            set
            {
                fid = value;
            }
        }


        #endregion

        #region ASSETS

        static string ipaddress;
        static string dns;
        static string netbios;
        static string operatingsystem;
        static string fqdn;

        public static string IPAddress
        {
            get
            {
                return ipaddress;
            }
            set
            {
                ipaddress = value;
            }
        }
        public static string Dns
        {
            get
            {
                return dns;
            }
            set
            {
                dns = value;
            }
        }
        public static string NetBios
        {
            get
            {
                return netbios;
            }
            set
            {
                netbios = value;
            }
        }
        public static string OperatingSystem
        {
            get
            {
                return operatingsystem;
            }
            set
            {
                operatingsystem = value;
            }
        }
        public static string FQDN
        {
            get
            {
                return fqdn;
            }
            set
            {
                fqdn = value;
            }
        }
        #endregion
        
        #region VULNERABILITY

        static int qid;
        static string title;
        static string vulntype;
        static int severity;
        static string cveid;
        static string vendorref;
        static string bugtraq;
        static string threat;
        static string impact;
        static string solution;
        static string exploitability;
        static string associatedmalware;
        static string pci;
        static string instance;
        static string category;

        public static int QID
        {
            get
            {
                return qid;
            }
            set
            {
                qid = value;
            }
        }


        public static string Title
        {
            get
            {
                return title;
            }
            set
            {
                title = value;
            }
        }

        public static string VulnerabilityType
        {
            get
            {
                return vulntype;
            }
            set
            {
                vulntype = value;
            }
        }

        public static int Severity
        {
            get
            {
                return severity;
            }
            set
            {
                severity = value;
            }
        }

        public static string CVE
        {
            get
            {
                return cveid;
            }
            set
            {
                cveid = value;
            }
        }

        public static string VendorReference
        {
            get
            {
                return vendorref;
            }
            set
            {
                vendorref = value;
            }
        }

        public static string BugTraq_ID
        {
            get
            {
                return bugtraq;
            }
            set
            {
                bugtraq = value;
            }
        }

        public static string Threat
        {
            get
            {
                return threat;
            }
            set
            {
                threat = value;
            }
        }

        public static string Impact
        {
            get
            {
                return impact;
            }
            set
            {
                impact = value;
            }
        }

        public static string Solution
        {
            get
            {
                return solution;
            }
            set
            {
                solution = value;
            }
        }

        public static string Exploitability
        {
            get
            {
                return exploitability;
            }
            set
            {
                exploitability = value;
            }
        }

        public static string AssociatedMalware
        {
            get
            {
                return associatedmalware;
            }
            set
            {
                associatedmalware = value;
            }
        }

        public static string PCI
        {
            get
            {
                return pci;
            }
            set
            {
                pci = value;
            }
        }

        public static string Instance
        {
            get
            {
                return instance;
            }
            set
            {
                instance = value;
            }
        }

        public static string Category
        {
            get
            {
                return category;
            }
            set
            {
                category = value;
            }
        }

        #endregion

        #region FINDINGS
        static string port;
        static string protocol;
        static string results;
        static string status;
        static string tempstring;
        static DateTime month;
        static DateTime year;

        public static string Port
        {
            get
            {
                return port;
            }
            set
            {
                port = value;
            }
        }

        public static string Protocol
        {
            get
            {
                return protocol;
            }
            set
            {
                protocol = value;
            }
        }

        public static string Results
        {
            get
            {
                return results;
            }
            set
            {
                results = value;
            }
        }

        public static string Status
        {
            get
            {
                return status;
            }
            set
            {
                status = value;
            }
        }

        public static string TempString
        {
            get
            {
                return tempstring;
            }
            set
            {
                tempstring = value;
            }
        }

        public static DateTime Month
        {
            get
            {
                return month;
            }
            set
            {
                month = value;
            }
        }

        public static DateTime Year
        {
            get
            {
                return year;
            }
            set
            {
                year = value;
            }
        }

        #endregion
    }
}
