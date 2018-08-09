using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace dsutil
{
    public class DSUtil
    {
        // Use this version of send when deploying
        public static async Task<string> SendMailProd(string emailTo, string body, string subject, string host)
        {
            string ret = null;
            try
            {
                MailMessage mailMessage = new MailMessage();
                MailAddress fromAddress = new MailAddress("onepluswonder@gmail.com");
                mailMessage.From = fromAddress;
                mailMessage.To.Add(emailTo);
                mailMessage.Body = body;
                mailMessage.IsBodyHtml = true;
                mailMessage.Subject = subject;
                SmtpClient smtpClient = new SmtpClient();
                smtpClient.Host = host;
                await smtpClient.SendMailAsync(mailMessage);
            }
            catch (Exception exc)
            {
                ret = exc.Message;
            }
            return ret;
        }

        // Use this version of send when working in development
        // Can try using just Send(), see if it works.
        public static async Task<string> SendMailDev(string toAddress, string subject, string body)
        {
            string ret = null;
            try
            {
                // Gmail Address from where you send the mail
                var fromAddress = "ventures2019@gmail.com";
                // any address where the email will be sending
                //Password of your gmail address
                const string fromPassword = "k3918834";
                // Passing the values and make a email formate to display
                // smtp settings
                var smtp = new System.Net.Mail.SmtpClient();
                {
                    smtp.Host = "smtp.gmail.com";
                    smtp.Port = 587;
                    smtp.EnableSsl = true;
                    smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                    smtp.Credentials = new NetworkCredential(fromAddress, fromPassword);
                    smtp.Timeout = 20000;
                }
                // Passing values to smtp object
                await smtp.SendMailAsync(fromAddress, toAddress, subject, body);
            }
            catch (Exception exc)
            {
                ret = exc.Message;
            }
            return ret;
        }

        public static void WriteFile(string filename, string msg, bool blankLine = false)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(System.AppDomain.CurrentDomain.BaseDirectory + filename, true))
            {
                if (!blankLine)
                {
                    string dtStr = DateTime.Now.ToString();
                    file.WriteLine(dtStr + " " + msg);
                }
                else
                {
                    file.WriteLine(msg);
                }
            }
        }
        // convert array of strings to delimited string
        public static string ListToDelimited(string[] array, char limiter)
        {
            string s = "";
            foreach (string i in array)
                s += i + limiter;
            if (s.Length > 0)
                s = s.Substring(0, s.Length - 1);   // remove last semi colon
            return s;
        }

        // convert delimited string to list of strings
        public static List<string> DelimitedToList(string str, char limiter)
        {
            string[] array = str.Split(limiter);
            return array.ToList();
        }
    }
}
