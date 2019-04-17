using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using PtacDealerExcelToTableService;
using Quartz;

using System.Data.OleDb;
using System.Data.SqlClient;
using Quartz.Impl;

namespace PtacDealerExcelToTableService
{
    public partial class PtacDealerExcelToTableService : ServiceBase
    {
        ISchedulerFactory schedFact;// = new StdSchedulerFactory();
        IScheduler sched;

        public PtacDealerExcelToTableService()
        {
            InitializeComponent();
        }

        //protected override void OnStart(string[] args)

        public void OnStart(string[] args)
        {
            string logPath = System.Configuration.ConfigurationManager.AppSettings["LoggerPath"].ToString(); //establish con. string to fetch log path//
            Logger.ConfigureLogger(logPath);

            string scheduleTime = System.Configuration.ConfigurationManager.AppSettings["Time"].ToString(); //schedules time how often the service to run//

            schedFact = new StdSchedulerFactory();
            sched = schedFact.GetScheduler();
            //IScheduler sched = schedFact.GetScheduler().GetAwaiter().GetResult();

            Logger.Info("........ Instantiate Schedular ........");
            IJobDetail transactionJob = JobBuilder.Create<PtacDealerExcelDataDBDump>()
                                                  .WithIdentity("TransactionProcessing", "InfoFinder")
                                                  .Build();
            Logger.Info("........ Instantiate Job ........");
            var trigger = TriggerBuilder.Create()
                          .WithIdentity("transactionTrigger", "InfoFinder")
                          .WithCronSchedule(scheduleTime.ToString())
                //.ForJob("TransactionJob")
                //.ForJob(transactionJob)
                          .Build();
            Logger.Info("........ Instantiate trigger ........");
            // sched.ScheduleJob(transactionJob, trigger);
            sched.ScheduleJob(transactionJob, trigger);


            sched.Start();
            Logger.Info("........ Scheduled job ........");
        }

        protected override void OnStop()
        {
        }
    } 


    public class PtacDealerExcelDataDBDump : IJob
    {
        static StringBuilder sb = new StringBuilder();

        public void ExcelToTableService()
        {
            // System.Diagnostics.Debugger.Launch();
            string filePath = "D:\\PTACDealerList.xlsx";

            if (System.IO.File.Exists(filePath))
            {
                string extension = Path.GetExtension(filePath);
                string conString = string.Empty;
                switch (extension)
                {
                    case ".xls": //Excel 97-03.
                        conString = System.Configuration.ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;

                        break;
                    case ".xlsx": //Excel 07 and above.
                        conString = System.Configuration.ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                        break;
                }


                DataTable dt = new DataTable();

                conString = string.Format(conString, filePath);

                using (OleDbConnection connExcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {
                            cmdExcel.Connection = connExcel;

                            //Get the name of First Sheet.
                            connExcel.Open();
                            DataTable dtExcelSchema;
                            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                            connExcel.Close();

                            //Read Data from First Sheet.
                            connExcel.Open();
                            cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                            odaExcel.SelectCommand = cmdExcel;
                            odaExcel.Fill(dt);
                            foreach (DataRow row in dt.Rows)
                            {
                                var Description = row["SUBFDESC"].ToString().Trim();
                                var Address = row["SUBFADR1"].ToString().Trim();
                                var City = row["SUBFCITY"].ToString().Trim();
                                var State = row["SUBFSTATE"].ToString().Trim();
                                var ZipCode = row["SUBFZIP"].ToString().Trim();
                                var PhoneNumber = row["PhoneNum"].ToString().Trim();
                                var PhoneNumber2 = row["PhoneNum"].ToString().Trim();

                                conString = System.Configuration.ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                                using (SqlConnection con = new SqlConnection(conString))
                                {
                                    String query = "INSERT INTO dbo.PtacDealer_TB (Description,Address,City,State,ZipCode,PhoneNumber,PhoneNumber2) VALUES ('" + Description.Replace("'", "''") + "','" + Address.Replace("'", "''") + "','" + City.Replace("'", "''") + "','" + State + "','" + ZipCode + "','" + PhoneNumber + "','" + PhoneNumber2 + "')";
                                    
                                    SqlCommand command = new SqlCommand(query, con);

                                    //var desc =  command.Parameters.Add("@Description", Description);
                                    //  command.Parameters.Add("@Address", Address);
                                    //  command.Parameters.Add("@City", City);
                                    //  command.Parameters.Add("@State", State);
                                    //  command.Parameters.Add("@ZipCode", ZipCode);
                                    //  command.Parameters.Add("@PhoneNumber", PhoneNumber);
                                    //  command.Parameters.Add("@PhoneNumber2", PhoneNumber2);
                                    //con.Open();
                                    //string  result = command.ExecuteNonQuery();


                                    con.Open();
                                    command.ExecuteNonQuery();
                                    con.Close();
                                    //command.ExecuteNonQuery();
                                    connExcel.Close();
                                }


                            }
                        }
                    }
                }
            }
        }
         
        public async void Execute(IJobExecutionContext context)
        {


            //-------calling Service to import Data from Excel to Table-----//
            ExcelToTableService();
            MailHelper mh = new MailHelper();
            try
            {
                Logger.Info("........ Process of moving dump started ........");
                #region //Extract and Move to newCSV folder

                string dataDumpZipPath = ConfigurationManager.AppSettings["dataDumpZipPath"].ToString();
                string zipExtractPath = ConfigurationManager.AppSettings["zipExtractPath"].ToString();

                string dbDumpSourceFile = zipExtractPath;
                string dbDumpDestination = ConfigurationManager.AppSettings["dbDumpDestination"].ToString();

                string filename = DateTime.Now.ToString("ddMMMMyyyy"); ;
                string dbDumpBackupDestination = ConfigurationManager.AppSettings["dbDumpBackupDestination"].ToString() + filename;

                sb.Append("<br/><h3><b>InfoFinder DB Dump, Image and Literature moving from FTP to Root </b></h3><br/><table border=\"1\"> <thead> <th>Process</th></thead>");
                //System.IO.Directory.CreateDirectory(backupDestination);

                try
                {
                    if (File.Exists(dataDumpZipPath))
                    {
                        string dbDumpDate = File.GetLastWriteTime(dataDumpZipPath).ToShortDateString();
                        if (System.IO.Directory.Exists(dbDumpBackupDestination))
                        {
                            Directory.Delete(dbDumpBackupDestination, true);
                            System.IO.Directory.Move(dbDumpDestination, dbDumpBackupDestination);
                            Logger.Info("Db dump backup done " + dbDumpBackupDestination);
                        }
                        else
                        {
                            System.IO.Directory.Move(dbDumpDestination, dbDumpBackupDestination);
                            Logger.Info("Db dump backup done " + dbDumpBackupDestination);
                        }

                        if (System.IO.Directory.Exists(zipExtractPath))
                        {
                            Directory.Delete(zipExtractPath, true);
                            ZipFile.ExtractToDirectory(dataDumpZipPath, zipExtractPath);
                            Logger.Info("Zip file extracted to " + zipExtractPath);
                        }
                        else
                        {
                            ZipFile.ExtractToDirectory(dataDumpZipPath, zipExtractPath);
                            Logger.Info("Zip file extracted to " + zipExtractPath);
                        }

                        System.IO.Directory.Move(dbDumpSourceFile, dbDumpDestination);
                        Logger.Info("Data dump moved from " + dbDumpSourceFile + " to " + dbDumpDestination);
                        //if (!System.IO.Directory.Exists(destinationFile))
                        //{
                        //    //System.IO.Directory.CreateDirectory(destinationFile);
                        //    System.IO.Directory.Move(sourceFile, destinationFile);
                        //}           

                        sb.Append("<tr><td>Moved Data dump dated <b>" + dbDumpDate + "</b> from <b>" + dataDumpZipPath + "</b> to <b>" + dbDumpDestination + "</b> folder</td></tr>");

                    }
                    else
                    {
                        sb.Append("<tr><td>Did not find the dump in the specified folder<b> " + dataDumpZipPath + "</b></td></tr>");
                    }
                    Logger.Info("........ Process of moving dump ended ........");
                }
                catch (Exception ex)
                {

                    sb.Append("<tr><td>Not able to move the dump due to following error <b>" + ex.Message + "</b></td></tr>");
                }

                #endregion

                # region // Pdf Copy


                string pdfTargetPath = ConfigurationManager.AppSettings["pdfTargetPath"].ToString();
                string pdfSourcePath = ConfigurationManager.AppSettings["pdfSourcePath"].ToString();

                try
                {
                    Logger.Info("........ Process of copying pdf's started........");
                    if (!System.IO.Directory.Exists(pdfTargetPath))
                    {
                        System.IO.Directory.CreateDirectory(pdfTargetPath);
                    }

                    if (System.IO.Directory.Exists(pdfSourcePath))
                    {
                        string[] files = System.IO.Directory.GetFiles(pdfSourcePath);

                        // Copy the files and overwrite destination files if they already exist.
                        foreach (string s in files)
                        {
                            string fileName = "";
                            string destFile = "";
                            try
                            {
                                // Use static Path methods to extract only the file name from the path.
                                fileName = System.IO.Path.GetFileName(s);
                                destFile = System.IO.Path.Combine(pdfTargetPath, fileName);
                                System.IO.File.Copy(s, destFile, true);
                            }
                            catch (Exception ex)
                            {
                                Logger.Info("Issue copying file" + fileName + "due to" + ex.Message);
                            }

                        }
                        sb.Append("<tr><td>Copied <b>" + files.Count().ToString() + "</b> pdf's from <b>" + pdfSourcePath + "</b> to <b>" + pdfTargetPath + " </b></td></tr>");
                    }
                    Logger.Info("........ Process of copying pdf's ended........");
                }
                catch (Exception ex)
                {
                    sb.Append("<tr><td>Not able to copy the pdf's due to following error <b>" + ex.Message + "</b></td></tr>");
                    Logger.Info(" Process of copying pdf's issue" + ex.Message);
                }
                //Directory.Delete(PdftargetPath, true);



                #endregion

                #region // Image Copy


                string imageTargetPath = ConfigurationManager.AppSettings["imageTargetPath"].ToString();
                string imageSourcePath = ConfigurationManager.AppSettings["imageSourcePath"].ToString();
                try
                {
                    Logger.Info("........ Process of copying images started........");
                    if (!System.IO.Directory.Exists(imageTargetPath))
                    {
                        System.IO.Directory.CreateDirectory(imageTargetPath);
                    }

                    if (System.IO.Directory.Exists(imageSourcePath))
                    {
                        string[] files = System.IO.Directory.GetFiles(imageSourcePath);

                        // Copy the files and overwrite destination files if they already exist.
                        foreach (string s in files)
                        {
                            string fileName = "";
                            string destFile = "";
                            try
                            {
                                fileName = System.IO.Path.GetFileName(s);
                                destFile = System.IO.Path.Combine(imageTargetPath, fileName);
                                System.IO.File.Copy(s, destFile, true);
                            }
                            catch (Exception ex)
                            {
                                Logger.Info("Issue copying file" + fileName + "due to" + ex.Message);

                            }
                            // Use static Path methods to extract only the file name from the path.

                        }
                        sb.Append("<tr><td>Copied <b>" + files.Count().ToString() + "</b> images from <b>" + imageSourcePath + "</b> to <b>" + imageTargetPath + "</b></td></tr></table><br/>");
                    }
                    Logger.Info("........ Process of copying images ended........");
                }
                catch (Exception ex)
                {
                    sb.Append("<tr><td>Not able to copy the images due to following error <b>" + ex.Message + "</b></td></tr></table><br/>");
                    Logger.Info(" Process of copying images issue" + ex.Message);
                }
                //Directory.Delete(ImagetargetPath, true);   
                #endregion
                mh.SendEmail(System.Configuration.ConfigurationManager.AppSettings["ReceiverEmail"].ToString(), "Infofinder Database Dump replacement Service ", "<h2>Report for the Date:" + DateTime.Today.AddDays(-1).ToShortDateString().ToString() + "</h2> <br/>" + sb.ToString());

            }
            catch (Exception ex)
            {
                sb.Clear();
                sb.Append("<h3>Error </h3><br/>Reason: " + ex.Message);
                mh.SendEmail(System.Configuration.ConfigurationManager.AppSettings["ReceiverEmail"].ToString(), "Infofinder Database Dump replacement error report  ", "<h2>Report for the Date:" + DateTime.Today.AddDays(-1).ToShortDateString().ToString() + "</h2> <br/>" + sb.ToString());
                Logger.Info("Error due to " + ex.Message);
            }

        }


        public static void InitiateLogger()
        {
            string logPath = System.Configuration.ConfigurationManager.AppSettings["LoggerPath"].ToString();
            Logger.ConfigureLogger(logPath);
        }
    }



}
