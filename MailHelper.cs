using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace PtacDealerExcelToTableService
{
    class MailHelper
    {
        public MailHelper()
        {         
            this.Host = ConfigurationManager.AppSettings["Host"].ToString();
            this.Port = Int32.Parse(ConfigurationManager.AppSettings["Port"]);
            this.Username = ConfigurationManager.AppSettings["EmailUsername"].ToString();
            this.Password = ConfigurationManager.AppSettings["EmailPassword"].ToString();
            this.From = ConfigurationManager.AppSettings["From"].ToString();
            this.EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableSsl"]);
        }
        public string From { get; set; }
        public string Host { get; set; }
        public int Port { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public bool EnableSsl { get; set; }

        public bool SendEmail(string Email, string Subject, string Body)
        {
            try
            {
                var smtpClient = new SmtpClient(this.Host, this.Port);
                smtpClient.EnableSsl = this.EnableSsl;
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = new NetworkCredential(this.Username, this.Password);
                var message =
                    new MailMessage(this.From, Email)
                    {
                        Subject = Subject,
                        Body = Body
                    };
                message.IsBodyHtml = true;
                smtpClient.Send(message);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }
}
