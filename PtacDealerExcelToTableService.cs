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
using System.Configuration;

namespace PtacDealerExcelToTableService
{
    public partial class PtacDealerExcelToTableService : ServiceBase
    {
        ISchedulerFactory schedFact;
        IScheduler sched;

        public PtacDealerExcelToTableService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            string logPath = ConfigurationManager.AppSettings["LoggerPath"].ToString(); //establish con. string to fetch log path//
            Logger.ConfigureLogger(logPath);
            Logger.Info("........ On start ........");
            string scheduleTime = ConfigurationManager.AppSettings["Time"].ToString(); //schedules time how often the service to run//

            schedFact = new StdSchedulerFactory();
            sched = schedFact.GetScheduler();

            Logger.Info("........ Instantiate Schedular ........");
            IJobDetail transactionJob = JobBuilder.Create<PtacDealerExcelDataDBDump>()
                                                  .WithIdentity("TransactionProcessing", "PtacDealer")
                                                  .Build();
            Logger.Info("........ Instantiate Job ........");
            var trigger = TriggerBuilder.Create()
                          .WithIdentity("transactionTrigger", "PtacDealer")
                          .WithCronSchedule(scheduleTime.ToString())
                          .Build();
            Logger.Info("........ Instantiate trigger ........");
            sched.ScheduleJob(transactionJob, trigger);


            sched.Start();
            Logger.Info("........ Scheduled job ........");
        }

        protected override void OnStop()
        {
        }
        public static void InitiateLogger()
        {
            string logPath = ConfigurationManager.AppSettings["LoggerPath"].ToString();
            Logger.ConfigureLogger(logPath);
        }
    }
}



