﻿using Quartz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;


namespace PtacDealerExcelToTableService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
        //    ServiceBase[] ServicesToRun;
        //    ServicesToRun = new ServiceBase[] 
        //    { 
        //        new PtacDealerExcelToTableService() 
        //    };
        //    ServiceBase.Run(ServicesToRun);


            //PtacDealerExcelToTableService service = new PtacDealerExcelToTableService();

            PtacDealerExcelDataDBDump service1 = new PtacDealerExcelDataDBDump();
            service1.ExcelToTableService();

            //PtacDealerExcelToTableService service = new PtacDealerExcelToTableService();
            //service.OnStart(test);

        }
    }
}