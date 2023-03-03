using Microsoft.Extensions.Configuration;
//using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System.IO.Compression;
using System.Net.Mail;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Configuration;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.IO;
using OfficeOpenXml.Core.ExcelPackage;
using System.Data.OleDb;
using System.Data.Common;
using System.Runtime.Intrinsics.X86;
using static System.Net.Mime.MediaTypeNames;
using AutoReportApplication;
using OpenQA.Selenium.DevTools;
using RPASuiteDataService.DBEntities;
using DbDataReaderMapper;
using Microsoft.EntityFrameworkCore;

namespace Utilities
{
    public class Utility
    {
        static String downloadDirectory = "";
        static String emailmessage = "";
        static String emailsubject = "";

        private readonly IConfiguration _configuration;
        public Utility(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        public static string getConnectionString()
        {
            var path = "";
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettingsCG.json", optional: true, reloadOnChange: true);

            IConfigurationRoot Configuration = builder.Build();

            var obj = Configuration.GetSection("EmailSettings").Get<EmailConfiguration>();
            return (Configuration.GetConnectionString("DefaultConnection")).ToString();
        }
        public static FileInfo AddDateTimeToFileName(string path)
        {
            var newFile = new DirectoryInfo(path).GetFiles().OrderByDescending(o => o.LastWriteTime).FirstOrDefault();
            var downloadfile = Path.GetFileNameWithoutExtension(newFile.ToString());
            var ext = Path.GetExtension(newFile.ToString());
            var SavePath = Path.GetDirectoryName(newFile.FullName.ToString());
            downloadfile = SavePath + "\\" + downloadfile + "__" + DateTime.Now.ToString("MM-dd-yyyy_hh_mm_ss") + ext;
            File.Move(newFile.FullName.ToString(), downloadfile);
            FileInfo fileInfo = new FileInfo(downloadfile);
            return fileInfo;
        }

        #region Common
        public static void ReadZip(string path, string extractpath, string extension)
        {
            string zipPath = path;
            string extractPath = extractpath;
            try
            {

                using (ZipArchive archive = ZipFile.OpenRead(zipPath))
                {
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        if (entry.FullName.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
                        {
                            // Gets the full path to ensure that relative segments are removed.
                            string destinationPath = Path.GetFullPath(Path.Combine(extractPath, entry.FullName));
                            entry.ExtractToFile(destinationPath, true);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                emailsubject = "Exception Occured : Reading ZIP File";
                emailmessage = ex.Message + " Line Number : " + Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' '))) + "\n\n File Path : " + zipPath + "." + extension;
                Email_Notification(emailmessage, emailsubject);
                throw;
            }
        }
        public static string[] GetFilesinDirectory(string Folderpath)
        {
            return Directory.GetFiles(Folderpath);
        }
        public static void ClearDirectory(string FolderPath)
        {
            var files = GetFilesinDirectory(FolderPath);
            foreach (var ele in files)
            {
                File.Delete(ele);
            }
        }
        public static void DeleteZipAfterExtraction(string FolderPath)
        {
            var ziptoDelete = "";
            try
            {
                var files = GetFilesinDirectory(FolderPath);
                foreach (var ele in files)
                {
                    if (ele.EndsWith(".zip"))
                    {
                        ziptoDelete = ele.ToString();
                        File.Delete(ele);
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                emailsubject = "Exception Occured : Deleting ZIP File";
                emailmessage = ex.Message + " Line Number : " + Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' '))) + "\n\n File Name : " + ziptoDelete;
                Email_Notification(emailmessage, emailsubject);
                throw;
            }
        }
        public static void Email_Notification(String message, String subject)
        {
            var builder = new ConfigurationBuilder()
                 .SetBasePath(Directory.GetCurrentDirectory())
                 .AddJsonFile("appsettingsCG.json", optional: true, reloadOnChange: true);

            IConfigurationRoot Configuration = builder.Build();
            var obj = Configuration.GetSection("EmailSettings").Get<EmailConfiguration>();


            //Email send
            MailMessage Mail = new MailMessage();
            Mail.From = new MailAddress(obj.FromEmail, obj.FromName);

            SmtpClient smtpClient = new SmtpClient(obj.SmtpServer);
            smtpClient.Timeout = 1000000;
            smtpClient.Port = obj.SmtpPort;
            smtpClient.EnableSsl = true;
            smtpClient.Credentials = new System.Net.NetworkCredential(obj.User, obj.Password);
            
            foreach (var recipient in obj.EmailRecipients)
            {
                Mail.To.Add(new MailAddress(recipient));
            }

            foreach (var ccRecipient in obj.CarbonCopy)
            {
                Mail.CC.Add(ccRecipient);
            }
            subject = subject.Replace('\r', ' ').Replace('\n', ' ');
            Mail.Subject = subject;
            Mail.Body = message;
            smtpClient.Send(Mail);
        }
        #endregion Common

        #region DB Related Methods
        public static DataTable DataTableForInsertionMUE(DataTable data, string ExcelColIndexes, string DBColNames)
        {
            DataTable table = new DataTable();
            string cols = DBColNames;
            try
            {
                var colsval = cols.Split(',');
                var colindexes = ExcelColIndexes.Split('|');

                for (int i = 0; i < colsval.Length; i++)
                {
                    table.Columns.Add(colsval[i]);
                }
                for (int rs = 0; rs < data.Rows.Count; rs++)
                {
                    DataRow _ravi = table.NewRow();
                    _ravi[colsval[0].ToString()] = data.Rows[rs][Convert.ToInt32(colindexes[0])].ToString();
                    _ravi[colsval[1].ToString()] = data.Rows[rs][Convert.ToInt32(colindexes[1])].ToString();
                    _ravi[colsval[2].ToString()] = data.Rows[rs][Convert.ToInt32(colindexes[2])].ToString(); ;
                    _ravi[colsval[3].ToString()] = data.Rows[rs][Convert.ToInt32(colindexes[3])].ToString();
                    table.Rows.Add(_ravi);
                }
                table.AcceptChanges();
                return table;
            }
            catch (Exception ex)
            {
                emailsubject = "Exception Occured : MUE Data table Insertion";
                emailmessage = ex.Message + " Line Number : " + Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')));
                Email_Notification(emailmessage, emailsubject);
                throw;
            }


        }
        public static void InsertMUEtoDB(FileInfo path)
        {

            var connStr = getConnectionString();
            var ColumnList = "";
            var Excelindexes = "";
            var ColumnNames = "";
            SqlConnection sqlConn = new SqlConnection(connStr);
            sqlConn.Open();
            SqlCommand cmd = new SqlCommand("Select ColumnList, ExcelIndexes,ColumnNames from [AscendTools].[RPASuite].[Datalistmapping] where PMSVersionListID = 9", sqlConn);
            SqlDataReader rdr;
            rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                ColumnList = rdr["ColumnList"].ToString();
                Excelindexes = rdr["Excelindexes"].ToString();
                ColumnNames = rdr["ColumnNames"].ToString();
            }
            try
            {
                //Excel connection strings for xls  
                string excelconnectionstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";

                //series of commands to bulk copy data from the excel file into our sql table   
                OleDbConnection connExcel = new OleDbConnection(excelconnectionstring);
                OleDbCommand cmdExcel = new OleDbCommand();
                OleDbDataAdapter oda = new OleDbDataAdapter();
                DataTable dt = new DataTable();
                cmdExcel.Connection = connExcel;
                // logerror("UploadPatFile", "Upload File", "1", "Pat 2");
                //Get the name of First Sheet

                connExcel.Open();
                //   logerror("UploadPatFile", "Upload File", "1", "Pat 3");
                DataTable dtExcelSchema;
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                int sheetCount = dtExcelSchema.Rows.Count;

                // int SheetNumber = int.Parse(Session["SheetNumber"].ToString());

                string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                connExcel.Close();

                //Read Data from First Sheet
                connExcel.Open();
                cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
                oda.SelectCommand = cmdExcel;
                oda.Fill(dt);
                connExcel.Close();
                DataTable processed = DataTableForInsertionMUE(dt, Excelindexes, ColumnNames);
                using (var connection = new SqlConnection(connStr))
                {
                    connection.Open();
                    //1- Delete existing data
                    using (SqlCommand command = new SqlCommand("DELETE FROM [AscendTools].[Staging].[MUE]", connection))
                    {
                        command.ExecuteNonQuery();
                    }
                    var transaction = connection.BeginTransaction();
                    using (var sqlBulk = new SqlBulkCopy(connection, SqlBulkCopyOptions.Default, transaction))
                    {
                        // SET BatchSize value.
                        sqlBulk.BatchSize = 500;
                        sqlBulk.DestinationTableName = "[AscendTools].[Staging].[MUE]";
                        sqlBulk.BulkCopyTimeout = 0;
                        sqlBulk.WriteToServer(processed);
                        transaction.Commit();
                    }
                }

            }
            catch (Exception ex)
            {
                emailsubject = "Exception Occured : Inserting MUE Data to DB";
                emailmessage = ex.Message + " Line Number : " + Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' '))) + "\n\n File Path : " + path;
                Email_Notification(emailmessage, emailsubject);
            }
        }
        public static DataTable DataTableForInsertionP2P(DataTable data, string ExcelColIndexes, string DBColNames)
        {
            DataTable table = new DataTable();
            string cols = DBColNames;
            try
            {
                var colsval = cols.Split(',');
                var colindexes = ExcelColIndexes.Split('|');

                for (int i = 0; i < colsval.Length; i++)
                {
                    table.Columns.Add(colsval[i]);
                }
                for (int rs = 0; rs < data.Rows.Count; rs++)
                {
                    DataRow _ravi = table.NewRow();
                    _ravi[colsval[0].ToString()] = data.Rows[rs][Convert.ToInt32(colindexes[0])].ToString();
                    _ravi[colsval[1].ToString()] = data.Rows[rs][Convert.ToInt32(colindexes[1])].ToString();
                    _ravi[colsval[2].ToString()] = data.Rows[rs][Convert.ToInt32(colindexes[2])].ToString();
                    _ravi[colsval[3].ToString()] = data.Rows[rs][Convert.ToInt32(colindexes[3])].ToString();
                    _ravi[colsval[4].ToString()] = data.Rows[rs][Convert.ToInt32(colindexes[4])].ToString();
                    _ravi[colsval[5].ToString()] = data.Rows[rs][Convert.ToInt32(colindexes[5])].ToString();
                    _ravi[colsval[6].ToString()] = data.Rows[rs][Convert.ToInt32(colindexes[6])].ToString();
                    table.Rows.Add(_ravi);
                }
                table.AcceptChanges();
                return table;
            }
            catch (Exception ex)
            {
                emailsubject = "Exception Occured : P2P Data table Insertion";
                emailmessage = ex.Message + " Line Number : " + Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')));
                Email_Notification(emailmessage, emailsubject);
                throw;
            }
        }
        public static void InsertP2PtoDB(FileInfo path)
        {
            var connStr = getConnectionString();
            var ColumnList = "";
            var Excelindexes = "";
            var ColumnNames = "";
            SqlConnection sqlConn = new SqlConnection(connStr);
            sqlConn.Open();
            SqlCommand cmd = new SqlCommand("Select ColumnList, ExcelIndexes,ColumnNames from [AscendTools].[RPASuite].[Datalistmapping] where PMSVersionListID = 11", sqlConn);
            SqlDataReader rdr;
            rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                ColumnList = rdr["ColumnList"].ToString();
                Excelindexes = rdr["Excelindexes"].ToString();
                ColumnNames = rdr["ColumnNames"].ToString();
            }
            try
            {
                //Excel connection strings for xls  
                string excelconnectionstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";

                //series of commands to bulk copy data from the excel file into our sql table   
                OleDbConnection connExcel = new OleDbConnection(excelconnectionstring);
                OleDbCommand cmdExcel = new OleDbCommand();
                OleDbDataAdapter oda = new OleDbDataAdapter();
                DataTable dt = new DataTable();
                cmdExcel.Connection = connExcel;
                // logerror("UploadPatFile", "Upload File", "1", "Pat 2");
                //Get the name of First Sheet

                connExcel.Open();
                //   logerror("UploadPatFile", "Upload File", "1", "Pat 3");
                DataTable dtExcelSchema;
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                int sheetCount = dtExcelSchema.Rows.Count;

                string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                connExcel.Close();

                //Read Data from First Sheet
                connExcel.Open();
                cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
                oda.SelectCommand = cmdExcel;
                oda.Fill(dt);
                connExcel.Close();
                DataTable processed = DataTableForInsertionP2P(dt, Excelindexes, ColumnNames);

                using (var connection = new SqlConnection(connStr))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("DELETE FROM [AscendTools].[Staging].[P2P]", connection))
                    {
                        command.ExecuteNonQuery();
                    }
                    var transaction = connection.BeginTransaction();
                    using (var sqlBulk = new SqlBulkCopy(connection, SqlBulkCopyOptions.Default, transaction))
                    {
                        // SET BatchSize value.
                        sqlBulk.BatchSize = 500;
                        sqlBulk.DestinationTableName = "[AscendTools].[Staging].[P2P]";
                        sqlBulk.BulkCopyTimeout = 0;
                        sqlBulk.WriteToServer(processed);
                        transaction.Commit();
                    }
                }               
            }
            catch (Exception ex)
            {
                emailsubject = "Exception Occured : Inserting P2P to DB";
                emailmessage = ex.Message + " Line Number : " + Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' '))) + "File Path : " + path;
                Email_Notification(emailmessage, emailsubject);
            }
        }
        public static void AOCEtextFiletoDB(string path)
        {
            string textFile = path;
            DataTable table = new DataTable();

            try
            {
                var connStr = getConnectionString();

                table.Columns.Add("F1");
                if (File.Exists(textFile))
                {
                    // Read a text file line by line.  
                    string[] lines = File.ReadAllLines(textFile);
                    foreach (string line in lines)
                    {
                        DataRow _ravi = table.NewRow();
                        _ravi["F1"] = line;
                        table.Rows.Add(_ravi);
                    }
                    var dt = table;
                    using (var con = new SqlConnection(connStr))
                    {
                        con.Open();
                        //1- Delete existing data
                        using (SqlCommand command = new SqlCommand("DELETE FROM [AscendTools].[Staging].[AddOnEdits]", con))
                        {
                            command.ExecuteNonQuery();
                        }
                        var transaction = con.BeginTransaction();
                        using (var sqlBulk = new SqlBulkCopy(con, SqlBulkCopyOptions.Default, transaction))
                        {
                            // SET BatchSize value.
                            sqlBulk.BatchSize = 500;
                            sqlBulk.DestinationTableName = "[AscendTools].[Staging].[AddOnEdits]";
                            sqlBulk.BulkCopyTimeout = 0;
                            sqlBulk.WriteToServer(dt);
                            transaction.Commit();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                emailsubject = "Exception Occured : AOCE text fiel to DB";
                emailmessage = ex.Message + " Line Number : " + Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' '))) + "\n\n File Path : " + path;
                Email_Notification(emailmessage, emailsubject);
                throw;
            }
        }
        public static void WriteToServerCG(DataTable dt, string table)
        {

            String connStr = getConnectionString();
            try
            {
                using (var con = new SqlConnection(connStr))
                {
                    con.Open();
                    //1- Delete existing data
                    using (SqlCommand command = new SqlCommand("DELETE FROM " + table + "", con))
                    {
                        command.ExecuteNonQuery();
                    }

                    var transaction = con.BeginTransaction();
                    using (var sqlBulk = new SqlBulkCopy(con, SqlBulkCopyOptions.Default, transaction))
                    {
                        // SET BatchSize value.
                        sqlBulk.BatchSize = 500;
                        sqlBulk.DestinationTableName = table;
                        sqlBulk.BulkCopyTimeout = 0;
                        sqlBulk.WriteToServer(dt);
                        transaction.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                emailsubject = "Exception Occured : Write to Server Claim Guard";
                emailmessage = ex.Message + " Line Number : " + Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')));
                throw;
            }
        }
        public static void CGUpdateLCDsAndArticlesFromMDBFile(string type, string filepath)
        {
            try
            {
                DataTable dtTableInfo = GetTablesInfo(type);
                var TablelistExisting = dtTableInfo.AsEnumerable().Select(r => r["TableName"].ToString());
                string[] Existing = TablelistExisting.ToArray();
                DbProviderFactories.RegisterFactory("System.Data.OleDb", System.Data.OleDb.OleDbFactory.Instance);

                //for Connection
                //var factory = DbProviderFactories.GetFactory("System.Data.SqlClient");
                //DbConnection connection = factory.CreateConnection();

                DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");
                using (DbConnection connection = factory.CreateConnection())
                {
                    connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath;

                    OleDbConnection connAccess = new OleDbConnection(connection.ConnectionString);
                    OleDbCommand cmdAccess = new OleDbCommand();
                    OleDbDataAdapter oda = new OleDbDataAdapter();
                    DataTable dt = new DataTable();
                    cmdAccess.Connection = connAccess;

                    connAccess.Open();
                    DataTable table = new DataTable();
                    string columns = "";
                    columns = "TableName,ColumnName";
                    String[] dataProperty = new String[250];


                    DataTable schema = connAccess.GetSchema("Columns");
                    var colsval = columns.Split(',');

                    for (int i = 0; i < colsval.Length; i++)
                    {
                        table.Columns.Add(colsval[i]);
                    }
                    for (int rs = 0; rs < schema.Rows.Count; rs++)
                    {
                        DataRow _ravi = table.NewRow();
                        _ravi[colsval[0]] = schema.Rows[rs][2].ToString();
                        _ravi[colsval[1]] = schema.Rows[rs][3].ToString();
                        table.Rows.Add(_ravi);
                    }
                    table.AcceptChanges();

                    var a = table.DefaultView.ToTable(true, "TableName");
                    var TablelistReceived = a.AsEnumerable().Select(r => r["TableName"].ToString());
                    string[] Received = TablelistReceived.ToArray();
                    string removedTables = string.Join(",", Existing.Except(Received));

                    using (OleDbConnection cnt = new OleDbConnection(connection.ConnectionString))
                    {
                        cnt.Open();
                        foreach (DataRow dr in a.Rows)
                        {
                            string tablename = "", columnslist = "";
                            int columnscount = 0;
                            if (dtTableInfo.Select("TableName = '" + dr["TableName"] + "'").Count() > 0)
                            {
                                tablename = dtTableInfo.Select("TableName = '" + dr["TableName"] + "'")[0][0].ToString();
                                columnscount = Convert.ToInt32(dtTableInfo.Select("TableName = '" + dr["TableName"] + "'")[0][1]);
                                columnslist = dtTableInfo.Select("TableName = '" + dr["TableName"] + "'")[0][2].ToString();

                                int colcountmdb = 0;
                                string columnsmdb = "", columnslistmdb = "";

                                DataTable dte = new DataTable();
                                string query = "SELECT * from " + dr["TableName"];
                                OleDbDataAdapter adapter = new OleDbDataAdapter(query, cnt);
                                DataSet ds = new DataSet();
                                adapter.Fill(ds);
                                if (ds.Tables[0].Rows.Count >= 0)
                                {
                                    dte = ds.Tables[0];
                                    colcountmdb = dte.Columns.Count;
                                    string[] AllColumnNames = (from DataColumn x
                                                      in dte.Columns.Cast<DataColumn>()
                                                               select x.ColumnName).ToArray();
                                    columnslistmdb = string.Join(",", AllColumnNames);

                                }
                                if (colcountmdb == columnscount)
                                {
                                    if (type == "articles")
                                        WriteToServerCG(dte, "[all_article].[dbo].[" + dr["TableName"].ToString() + "]");

                                    else
                                        WriteToServerCG(dte, "[all_lcd].[dbo].[" + dr["TableName"].ToString() + "]");

                                    Console.WriteLine("Updated data of Table: {0}", dr["TableName"]);
                                }

                                else
                                {
                                    var mg = "No of Columns not matching Table: " + dr["TableName"] + "";
                                    if (type == "articles")
                                    {
                                        emailmessage = "Claim Guard All Articles Data Update, No of Columns not matching of Table:" + dr["TableName"] + ". Columns New:  " + columnslistmdb + ". Columns Old: " + columnslist;
                                        emailsubject = "Claim Guard Articles Data Error";

                                        Utility.Email_Notification(emailmessage, emailsubject);
                                    }
                                    else
                                    {
                                        emailmessage = "Claim Guard All LCDs Data Update, No of Columns not matching of Table:" + dr["TableName"] + ".";
                                        emailsubject = "Claim Guard LCDs Data Error";

                                        Utility.Email_Notification(emailmessage, emailsubject);
                                    }
                                    // var res = LogError("CG All Articles Data Mapping", "CG Data Update", "CG Data Update", mg, -1);
                                }
                                //    Console.WriteLine("No of Columns not matching");

                            }
                            else
                            {
                                //   var msg = "Table: " + dr["TableName"] + " Does not Exists";
                                // var res = LogError("CG All Articles Data Mapping", "CG Data Update", "CG Data Update", msg, -1);

                                if (type == "articles")
                                {
                                    emailmessage = "Claim Guard All Articles Data Update " + "Table: " + dr["TableName"] + " Doesn't Exist";
                                    emailsubject = "Claim Guard Articles Data Error";

                                    Utility.Email_Notification(emailmessage, emailsubject);
                                }
                                else
                                {
                                    emailmessage = "Claim Guard All LCDs Data Update " + "Table: " + dr["TableName"] + " Doesn't Exist";
                                    emailsubject = "Claim Guard LCDs Data Error";

                                    Utility.Email_Notification(emailmessage, emailsubject);
                                }
                            }
                        }

                    }
                    if (removedTables.Length > 0)
                    {
                        if (type == "articles")
                        {
                            emailmessage = "Claim Guard All Articles Data Update. Tables (" + removedTables + ") do not exist in new DataBase.";
                            emailsubject = "Claim Guard Articles Database Error";

                            Utility.Email_Notification(emailmessage, emailsubject);
                        }
                        else
                        {
                            emailmessage = "Claim Guard All LCDs Data Update. Tables (" + removedTables + ") do not exist in new DataBase.";
                            emailsubject = "Claim Guard LCDs Database Error";

                            Utility.Email_Notification(emailmessage, emailsubject);
                        }
                    }
                    connAccess.Close();

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + " Line Number : " + Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' '))));
            }

        }
        public static DataTable GetTablesInfo(string TYPE)
        {
            // int batchno = 0;
            try
            {
                using (var connection = new SqlConnection(getConnectionString()))
                {
                    string query = "";
                    if (TYPE == "articles")
                        query = "[all_article].[dbo].[usp_GetTableList]";
                    else
                        query = "[all_lcd].[dbo].[usp_GetTableList]";

                    System.Data.DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter(query, connection);
                    da.Fill(dt);
                    return dt;
                }
            }
            catch (Exception ex)
            {
                emailsubject = "Exception Occured : Get Data Tables Info ";
                emailmessage = ex.Message + " Line Number : " + Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')));
                Email_Notification(emailmessage, emailsubject);
                throw;
            }

        }
        #endregion DB Related Methods

        #region Stored Procedures
        public static List<DataListProcessing> GetDataListForProcessing(int dataListID = 0, string pmsCode = "")
        {
            string connStr = getConnectionString();
            DbDataReader reader = null;
            try
            {
                var result = new List<DataListProcessing>();
                SqlConnection sqlConnection = new SqlConnection(connStr);
                sqlConnection.Open();
                var cmd = sqlConnection.CreateCommand();
                cmd.CommandText = "[RPASuite].[SP_DataListForProcessing]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter { ParameterName = "@DataListID", Value = (dataListID == 0) ? DBNull.Value : dataListID });
                cmd.Parameters.Add(new SqlParameter { ParameterName = "PMSCode", Value = (String.IsNullOrEmpty(pmsCode)) ? DBNull.Value : pmsCode });
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(reader.MapToObject<DataListProcessing>());
                }
                reader.Close();
                sqlConnection.Close();
                return result;
            }
            catch (Exception ex)
            {
                emailsubject = "Exception Occured : Get Data List for Processing";
                emailmessage = ex.Message + " Line Number : " + Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')));
                Email_Notification(emailmessage, emailsubject);
                throw;
            }
            finally
            {
                if (reader?.IsClosed == false)
                {
                    reader?.Close();
                }
            }
        }
        public static List<LocatorProcessing> GetLocatorMappingForProcessing(int dataListMappingID, string pmsCode)
        {
            string connStr = getConnectionString();
            DbDataReader reader = null;
            try
            {
                var result = new List<LocatorProcessing>();
                SqlConnection sqlConnection = new SqlConnection(connStr);
                sqlConnection.Open();
                var cmd = sqlConnection.CreateCommand();
                cmd.CommandText = "[RPASuite].[SP_LocatorMappingForProcessing]";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter { ParameterName = "@DataListMappingID", Value = dataListMappingID });
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(reader.MapToObject<LocatorProcessing>());
                }
                reader.Close();
                sqlConnection.Close();
                return result;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (reader?.IsClosed == false)
                {
                    reader?.Close();
                }
            }
        }
        #endregion Stored Procedures
        public static string DownloadReport(int datalistID, int datalistmappingId)
        {
            var datalist = GetDataListForProcessing(datalistID);
            foreach (DataListProcessing item in datalist)
            {
                if (item.DataListMappingID == datalistmappingId)
                {
                    downloadDirectory = item.Report_Path;
                    Directory.CreateDirectory(downloadDirectory);
                    var chromeOptions = new ChromeOptions();
                    ClearDirectory(downloadDirectory);

                    chromeOptions.AddUserProfilePreference("download.prompt_for_download", false);
                    chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
                    chromeOptions.AddUserProfilePreference("download.default_directory", downloadDirectory);

                    // set browser headleass
                    chromeOptions.AddArguments("--headless");
                    chromeOptions.AddArguments("--disable-gpu");
                    chromeOptions.AddArguments("--window-size=1280,800");
                    chromeOptions.AddArguments("--allow-insecure-localhost");

                    var chromeDriverService = ChromeDriverService.CreateDefaultService();
                    chromeDriverService.HideCommandPromptWindow = true;    // to hide the console.

                    IWebDriver driver = new ChromeDriver(chromeDriverService, chromeOptions);
                    driver.Manage().Window.Maximize();
                    try
                    {
                        driver.Navigate().GoToUrl(item.URL_Path);
                        Thread.Sleep(5000);
                        Console.WriteLine(item.VersionNumber + " Download started");
                        List<LocatorProcessing> locators = GetLocatorMappingForProcessing(datalistmappingId, string.Empty);
                        foreach (var locator in locators)
                        {
                            if (locator.Event == "Click")
                            {
                                driver.FindElement(By.XPath(locator.Locator)).Click();
                                Thread.Sleep(locator.WaitTime);
                            }
                        }
                        while (Directory.GetFiles(downloadDirectory).Count(i => i.EndsWith(".crdownload")) > 0)
                        {
                            Thread.Sleep(2000);
                        }
                        Thread.Sleep(20000);
                    }
                    catch (Exception ex)
                    {
                        //Send Error email notification
                        emailsubject = "Exception Occured : Downloading Report " + item.VersionNumber;
                        emailmessage = ex.Message + " Line Number : " + Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')));
                        Email_Notification(emailmessage, emailsubject);
                        driver.Quit();
                    }
                    finally
                    {
                        driver.Quit();
                    }
                }
            }

            return downloadDirectory;
        }
    }
}


