﻿using HtmlAgilityPack;
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
        //
        public static async Task<string> SendMailProdAsync(string emailTo, string body, string subject, string host)
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
                SmtpClient smtpClient = new SmtpClient(host);
                // smtpClient.Host = host;
                await smtpClient.SendMailAsync(mailMessage);
            }
            catch (Exception exc)
            {
                ret = exc.Message;
            }
            return ret;
        }

        public static string SendMailProd(string emailTo, string body, string subject, string host)
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
                // smtpClient.Host = host;
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.PickupDirectoryFromIis;
                smtpClient.Send(mailMessage);
            }
            catch (Exception exc)
            {
                ret = exc.Message;
            }
            return ret;
        }

        // Use this version of send when working in development
        // Can try using just Send(), see if it works.
        //
        // YOU MUST follow this link to work with a give gmail address:
        // https://stackoverflow.com/questions/18503333/the-smtp-server-requires-a-secure-connection-or-the-client-was-not-authenticated
        // Look at second response.  First logon to gmail with the account you want to gmail from and then click the link he
        // shows to turn allow Less Secure Sign-In.
        //
        public static string SendMailDev(string toEmail, string subject, string body)
        {
            string ret = null;
            try
            {
                // Gmail Address from where you send the mail
                //var fromAddress = "ventures2019@gmail.com";
                // any address where the email will be sending
                //Password of your gmail address
                const string fromPassword = "k3918834";
                var fromAddress = new MailAddress("ventures2019@gmail.com", "ventures2019");
                var toAddress = new MailAddress(toEmail, toEmail);
                // Passing the values and make a email formate to display
                // smtp settings
                var smtp = new System.Net.Mail.SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword),
                    Timeout = 20000
                };
                using (var message = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = true
                })
                {
                    smtp.Send(message);
                }
                // Passing values to smtp object
                //await smtp.SendMailAsync(fromAddress, toAddress, subject, body);
            }
            catch (Exception exc)
            {
                ret = exc.Message;
            }
            return ret;
        }

        public static void WriteFile(string filename, string msg, string username, bool blankLine = false)
        {
            try
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(System.AppDomain.CurrentDomain.BaseDirectory + filename, true))
                {
                    if (!blankLine)
                    {
                        string dtStr = DateTime.Now.ToString();
                        file.WriteLine(dtStr + " " + username + " " + msg);
                    }
                    else
                    {
                        file.WriteLine(msg);
                    }
                }
            }
            catch
            {

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

        public static string ErrMsg(string header, Exception exc)
        {
            string msg = header + " " + exc.Message;
            if (exc.InnerException != null)
            {
                msg += " " + exc.InnerException.Message;
                if (exc.InnerException.InnerException != null)
                {
                    msg += " " + exc.InnerException.InnerException.Message;
                }
            }
            return msg;
        }
        public static string HTMLToString(string html)
        {
            // https://stackoverflow.com/questions/4182594/grab-all-text-from-html-with-html-agility-pack
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);
            string output = null;
            foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//text()"))
            {
                output += node.InnerText;
            }

            output = doc.DocumentNode.SelectSingleNode("//body").InnerText;
            return output;
        }
    }
}
