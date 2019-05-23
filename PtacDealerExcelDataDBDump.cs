using Quartz;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PtacDealerExcelToTableService
{
    class PtacDealerExcelDataDBDump : IJob
    {
        //method to import data from Excel to Database by stored procedure//
        public int ExcelToTableService()
        {
            string fileName = ConfigurationManager.AppSettings["ExcelFolderPath"].ToString();

            List<PtacDealersModel> resultantData = ExcelUtility.retrieveExcelData(fileName);

            var paramsData = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("@dealersXml",resultantData.ToXML())
                };

            var storedProcedureName = ConfigurationManager.AppSettings["storedProcedureName"];
            SqlUtility.exeuteStorProc(storedProcedureName, paramsData);

            return resultantData.Count();
        }

        public async void Execute(IJobExecutionContext context)
        {
            StringBuilder sb = new StringBuilder();
            MailHelper mh = new MailHelper();

            //-------calling Service to import Data from Excel to Table-----//
            int dealerRecords = ExcelToTableService();

            try
            {
                mh.SendEmail(ConfigurationManager.AppSettings["ReceiverEmail"].ToString(),
                "Ptac Dealer's Excel Data dump Service ",
                "<h2>Report for the Date:" + DateTime.Today.AddDays(-1).ToShortDateString().ToString() +
                "</h2> <br/>" + sb.ToString() + "Amana Ptac Dealer's data import successful." +
                "<h3> Amana Ptac Dealer's Records Sync: </h3>" + dealerRecords);
            }
            catch (Exception ex)
            {
                sb.Clear();
                sb.Append("<h3>Error </h3><br/>Reason: " + ex.Message);
                mh.SendEmail(ConfigurationManager.AppSettings["ReceiverEmail"].ToString(),
               "Ptac Dealer's Excel Data dump error report  ",
               "<h2>Report for the Date:" +
                DateTime.Today.AddDays(-1).ToShortDateString().ToString() +
                "</h2> <br/>" + sb.ToString() + " Ptac Dealer's data import not successful." +
                "<h2> Amana Ptac Dealer's failed to Sync Records");
                Logger.Info("Error due to " + ex.Message);
            }
        }

    }

}
