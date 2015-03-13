using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;

namespace DailyNCSConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            //make sure the temporary directory exists
            string tempFileLoc = Environment.CurrentDirectory + "\\Resources\\tempfiles\\";
            bool isExists = System.IO.Directory.Exists(tempFileLoc);
            if (!isExists)
                System.IO.Directory.CreateDirectory(tempFileLoc);
                //if it exists, delete yesterday's file
            else
            {
                DirectoryInfo theDir = new DirectoryInfo(tempFileLoc);
                foreach (System.IO.FileInfo file in theDir.GetFiles()) file.Delete();
            }
            //create the datasets, name is important as that sets formatting later in EPHelper
            DataSet dtPrinter = new DataSet();
            DataSet dtCopier = new DataSet();

            //located below are the SQL stored procedure names that generate the pivot table displayed in the application
            dtPrinter.Tables.Add(dtCreate("QBS", "##dpttempt", "NCSGen", "QBS Printer"));
            dtCopier.Tables.Add(dtCreate("QBS", "##dpttempt", "NCSGenCopier", "QBS Copier"));
            dtPrinter.Tables.Add(dtCreate("CTX", "##dpttempt", "NCSGen", "CTX Printer"));
            dtCopier.Tables.Add(dtCreate("CTX", "##dpttempt", "NCSGenCopier", "CTX Copier"));
            dtPrinter.Tables.Add(dtCreate("", "##dpttemptTot", "NCSTotalGen", "Total Printer"));
            dtCopier.Tables.Add(dtCreate("", "##dpttemptTot", "NCSTotalGenCopier", "Total Copier"));
            //create the filename here, as we will be passing that to emailer
            string fileName = "DailyNCS-" + DateTime.Now.ToString("MM") + "-" + DateTime.Now.ToString("dd") + "-" + DateTime.Now.ToString("yy");
            EPHelper.GenerateExcel(dtPrinter, dtCopier, fileName);
            DailyNCSConsole.emailer.makeEmail(fileName);
        }
     
        private static string connString = DailyNCSConsole.NCSSettings.Default.SQLConString;

        //public variable for SQL Connecion string. ^ All attributes for connection string above are taken from the applications settings file.
        public static string conString { get { return connString; } }

        public static DataTable dtCreate(string region, string theTable, string theSP, string theName)
        {
            var dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(conString))
                {
                    using (SqlCommand cmd = new SqlCommand(theSP, con))
                    using (var da = new SqlDataAdapter(cmd))
                    {
                        //check to see if the table requires a region passed as a parameter
                        cmd.CommandType = CommandType.StoredProcedure;
                        if (theSP == "NCSGen" || theSP == "NCSGenCopier")
                        {
                            cmd.Parameters.AddWithValue("@reg", region);
                        }
                     
                        con.Open();
                        cmd.ExecuteNonQuery();
                        da.Fill(dt);

                    }
           
                    con.Close();

                }
            }
            catch (Exception X)
            {
                EventLog log = new EventLog();
                log.Source = "DailyNCSConsole";
                log.WriteEntry("A SQL Error has occurred: " + X, EventLogEntryType.Error);
            }
            dt.TableName = theName;
            return dt;
        }

    }
}

