using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Configuration;

namespace ConsoleApplication1
{
    class Config
    {
        static string dbname;
        static string dbhost;
        static string dbuser;
        static string dbpass;
        static string dbport;
        static int regionID;
        static int buID;
        static string rawdir;
        static string converteddir;
        static int filecount;
        static string outputdir;

        public static string dbName
        {
            get
            {
                return dbname;
            }
            set
            {
                dbname = value;
            }
        }

        public static string dbHost
        {
            get
            {
                return dbhost;
            }
            set
            {
                dbhost = value;
            }
        }

        public static string dbUser
        {
            get
            {
                return dbuser;
            }
            set
            {
                dbuser = value;
            }
        }

        public static string dbPass
        {
            get
            {
                return dbpass;
            }
            set
            {
                dbpass = value;
            }
        }

        public static string dbPort
        {
            get
            {
                return dbport;
            }
            set
            {
                dbport = value;
            }
        }

        public static int RegionID
        {
            get
            {
                return regionID;
            }
            set
            {
                regionID = value;
            }
        }

        public static int BuID
        {
            get
            {
                return buID;
            }
            set
            {
                buID = value;
            }
        }

        public static string RawDir
        {
            get
            {
                return rawdir;
            }
            set
            {
                rawdir = value;
            }
        }

        public static string ConvertedDir
        {
            get
            {
                return converteddir;
            }
            set
            {
                converteddir = value;
            }
        }

        public static int FileCount
        {
            get
            {
                return filecount;
            }
            set
            {
                filecount = value;
            }
        }

        public static string OutputDir
        {
            get
            {
                return outputdir;
            }
            set
            {
                outputdir = value;
            }
        }

    }
}
