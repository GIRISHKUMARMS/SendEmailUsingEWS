using System;
using System.Configuration;
using System.Net;
using Microsoft.Exchange.WebServices.Data;

namespace SendEmailFromEWS
{
    class Program
    {
        static void Main(string[] args)
        {
            string MailUser = ConfigurationManager.AppSettings["MailUser"];
            string MailPass = ConfigurationManager.AppSettings["mailPass"].ToString();
            string MailTo = ConfigurationManager.AppSettings["MailTo"].ToString();
            var isSentEmail = SendMail(MailUser, MailPass, MailTo);
        }
        public static bool SendMail(string MailUser, string MailPass, string MailTo)
        {
            try
            {
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                service.Credentials = new NetworkCredential(MailUser, MailPass);
                service.AutodiscoverUrl(MailUser, (a) =>
                {
                    return true;
                });
                // service.AutodiscoverUrl(MailUser);

                EmailMessage emailMessage = new EmailMessage(service);
                emailMessage.Subject = "Profanity Alert! | Profanity Found | Review";
                emailMessage.Body = new MessageBody("<table><tr><td>Hi Admin,</td><td></td><td></td></tr><tr><td></td><td></td><td></td></tr><tr><td></td><td></td><td></td></tr><tr><td>Name</td><td>:</td><td>Info.name</td></tr><tr><td></td><td></td><td></td></tr><tr><td>Email</td><td>:</td><td>info.EmailId </td></tr><tr><td></td><td></td><td></td></tr><tr><td>Details</td><td>:</td><td> info.Details </td></tr><tr><td></td><td></td><td></td></tr><tr><td>Date</td><td>:</td><td>" + DateTime.Now.ToString() + "</td></tr></table>");
                emailMessage.ToRecipients.Add(MailTo);
                emailMessage.SendAndSaveCopy();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }
    }
}
