using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Globalization;
using System.Diagnostics;
using System.IO;

namespace DailyNCSConsole
{

    class emailer
    {
        public static void makeEmail(string fileN)
        {


            string userEmail = NCSSettings.Default.ToEmail;

            string tehDay = String.Format("{0:D}", DateTime.Now);

            string emailTitle = "Daily NCS Report for " + tehDay;

            string tehFile = Environment.CurrentDirectory + "\\Resources\\tempfiles\\" + fileN + ".xlsx";

            LinkedResource sigPic = new LinkedResource(Environment.CurrentDirectory + "\\Resources\\sigLogo.jpg");
            sigPic.ContentId = "sigLogo";

            try
            {
                string body = "<p>Greetings!</p><p></p><p><pre>  </pre><dd>Attached is the Daily NCS Report for " + tehDay + "</dd><br><br>Thanks,<br><br><dd> - NWG IT Staff </dd><br><img src=cid:sigLogo>";
                MailMessage userMessage = new MailMessage();
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(body, null, "text/html");
                htmlView.LinkedResources.Add(sigPic);
                Attachment filAttachment = new Attachment(tehFile);
                userMessage.Attachments.Add(filAttachment);
                userMessage.AlternateViews.Add(htmlView);
                userMessage.Subject = emailTitle;

                userMessage.To.Add(userEmail);
                userMessage.From = new MailAddress(NCSSettings.Default.FromEmail);
                userMessage.IsBodyHtml = true;

                SmtpClient sCli = new SmtpClient(NCSSettings.Default.SMTPServer);
                sCli.Port = 25;
                sCli.Send(userMessage);
                filAttachment.Dispose();
                ////File.Delete(tehFile);
            }
            catch (Exception x)
            {
               
                EventLog log = new EventLog();
                log.Source = "DailyNCS";
                log.WriteEntry(x.Message, EventLogEntryType.Error);
            }

        }
    }
}
