﻿using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace dsutil
{
    public class DSUtil
    {
        readonly static string _logfile = "log.txt";

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
            string msg = header + " ERROR " + exc.Message;
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

            output = doc.DocumentNode.SelectSingleNode("//body").InnerText;
            return output;
        }
        public static string HTMLToString_Full(string html)
        {
            // https://stackoverflow.com/questions/4182594/grab-all-text-from-html-with-html-agility-pack
            HtmlDocument doc = new HtmlDocument();
            string output = null;
            if (!string.IsNullOrEmpty(html))
            {
                doc.LoadHtml(html);
                foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//text()"))
                {
                    output += node.InnerText;
                }
            }
            return output;
        }

        public static int FindError(string filename, string marker)
        {
            int match = 0;
            if (File.Exists(filename))
            {
                foreach (string line in File.ReadLines(filename))
                {
                    if (line.Contains(marker.ToLower()) || line.Contains(marker.ToUpper()))
                    {
                        ++match;
                    }
                }
            }
            return match;
        }
        public static string GetLastError(string filename, string marker)
        {
            string lastErr = null;
            if (File.Exists(filename))
            {
                foreach (string line in File.ReadLines(filename))
                {
                    if (line.Contains(marker.ToLower()) || line.Contains(marker.ToUpper()))
                    {
                        lastErr = line;
                    }
                }
            }
            return lastErr;
        }
        public static DateTime AddBusinessDays(DateTime date, int days)
        {
            if (days < 0)
            {
                throw new ArgumentException("days cannot be negative", "days");
            }

            if (days == 0) return date;

            if (date.DayOfWeek == DayOfWeek.Saturday)
            {
                date = date.AddDays(2);
                days -= 1;
            }
            else if (date.DayOfWeek == DayOfWeek.Sunday)
            {
                date = date.AddDays(1);
                days -= 1;
            }

            date = date.AddDays(days / 5 * 7);
            int extraDays = days % 5;

            if ((int)date.DayOfWeek + extraDays > 5)
            {
                extraDays += 2;
            }

            return date.AddDays(extraDays);
        }
        public static int GetBusinessDays(DateTime start, DateTime end)
        {
            if (start.DayOfWeek == DayOfWeek.Saturday)
            {
                start = start.AddDays(2);
            }
            else if (start.DayOfWeek == DayOfWeek.Sunday)
            {
                start = start.AddDays(1);
            }

            if (end.DayOfWeek == DayOfWeek.Saturday)
            {
                end = end.AddDays(-1);
            }
            else if (end.DayOfWeek == DayOfWeek.Sunday)
            {
                end = end.AddDays(-2);
            }

            int diff = (int)end.Subtract(start).TotalDays;

            int result = diff / 7 * 5 + diff % 7;

            if (end.DayOfWeek < start.DayOfWeek)
            {
                return result - 2;
            }
            else
            {
                return result;
            }
        }

        /// <summary>
        /// Probably should but this in walmart library - they have funny habit of placing question marks in odd places.
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool ContainsQuestionMark(string str, out string segment)
        {
            const string marker = "&#39;";
            const int segmentLength = 10;
            segment = null;
            bool fail = false;
            bool done = false;
            int pos = 0;

            if (string.IsNullOrEmpty(str))
            {
                return false;
            }
            string justText = DSUtil.HTMLToString_Full(str);
            do
            {
                pos = justText.IndexOf(marker, pos);
                if (pos > -1)
                {
                    if (pos < justText.Length)
                    {
                        char c = justText[pos + marker.Length];
                        fail = Char.IsLetterOrDigit(c);
                        if (fail)
                        {
                            // get segment that contains the question mark
                            if (pos > segmentLength)
                            {
                                segment = justText.Substring(pos - segmentLength, segmentLength);
                            }
                            else
                            {
                                segment = justText.Substring(0, pos);
                            }
                            if (pos + segmentLength < justText.Length)
                            {
                                var endSegment = justText.Substring(pos, segmentLength);
                                segment += endSegment;
                            }
                            else
                            {
                                segment += justText.Substring(pos, 1);
                            }
                            done = true;
                        }
                        else
                        {
                            pos += 2;
                        }
                    }
                    else
                    {
                        done = true;
                    }
                }
                else
                {
                    done = true;
                }
            } while (!done);
            return fail;
        }

        /// <summary>
        /// Red flag if supplier item description contains certain key words.
        /// Hard-coding for now - let's see where this goes.
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool ContationsKeyWords(string str, out List<string> help)
        {
            help = new List<string>();
            bool ret = false;
            if (string.IsNullOrEmpty(str))
            {
                return false;
            }
            string justText = DSUtil.HTMLToString_Full(str);
            int pos = justText.ToUpper().IndexOf("COMMENTS");
            if (pos > -1)
            {
                help.Add("contains COMMENTS");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("QUESTIONS");
            if (pos > -1)
            {
                help.Add("contains QUESTIONS");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("WALMART");
            if (pos > -1)
            {
                help.Add("contains WALMART");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("WARRANTY");
            if (pos > -1)
            {
                help.Add("contains WARRANTY");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("PACK");
            if (pos > -1)
            {
                help.Add("contains PACK");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("WARNING");
            if (pos > -1)
            {
                help.Add("contains WARNING");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("?");
            if (pos > -1)
            {
                help.Add("contains ?");
                ret = true;
            }
            return ret;
        }
        public static bool ContationsDisclaimer(string str)
        {
            const string marker = "We aim to show you accurate product information.";
            bool ret = false;
            if (string.IsNullOrEmpty(str))
            {
                return false;
            }
            int pos = str.ToUpper().IndexOf(marker.ToUpper());
            if (pos > -1)
            {
                ret = true;
            }
            return ret;
        }

        /// <summary>
        /// Despite using selenium, google eventually starts asking if you are a robot
        /// </summary>
        /// <param name="search"></param>
        /// <returns></returns>
        public static string GoogleSearchSelenium(string search)
        {
            string retURL = null;
            try
            {
                IWebDriver driver = new ChromeDriver();
                driver.Navigate().GoToUrl("https://www.google.com");
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(60);

                Thread.Sleep(2000);
                driver.FindElement(By.Name("q")).SendKeys(search);
                driver.FindElement(By.Name("q")).Submit();

                Thread.Sleep(1000);

                foreach (var item in driver.FindElements(By.TagName("a")))
                {
                    var x = item.GetAttribute("href");
                    if (x != null)
                    {
                        if (x.StartsWith("https://www.walmart.com"))
                        {
                            retURL = x;
                            break;
                        }
                    }
                }
                driver.Quit();
            }
            catch (Exception exc)
            {
                string header = "GoogleSearch";
                string msg = dsutil.DSUtil.ErrMsg(header, exc);
                dsutil.DSUtil.WriteFile(_logfile, msg, "");
            }
            return retURL;
        }

        /// <summary>
        /// Will eventuall return error 429
        /// </summary>
        /// <param name="keywords"></param>
        /// <param name="limit">must find match within first 'limit' results</param>
        /// <returns></returns>
        public static string GoogleSearch(string keywords, int limit)
        {
            string supplierURL = null;
            StringBuilder sb = new StringBuilder();
            byte[] ResultsBuffer = new byte[8192];
            string SearchResults = "http://google.com/search?q=" + keywords;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(SearchResults);

            // error is here: The remote server returned an error: (429) Too Many Requests.
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            Stream resStream = response.GetResponseStream();
            string tempString = null;
            int count = 0;
            do
            {
                count = resStream.Read(ResultsBuffer, 0, ResultsBuffer.Length);
                if (count != 0)
                {
                    tempString = Encoding.ASCII.GetString(ResultsBuffer, 0, count);
                    sb.Append(tempString);
                }
            }

            while (count > 0);
            string sbb = sb.ToString();

            HtmlAgilityPack.HtmlDocument html = new HtmlAgilityPack.HtmlDocument();
            html.OptionOutputAsXml = true;
            html.LoadHtml(sbb);
            HtmlNode doc = html.DocumentNode;

            int i = 0;
            var links = new List<string>();
            foreach (HtmlNode link in doc.SelectNodes("//a[@href]"))
            {
                //HtmlAttribute att = link.Attributes["href"];
                string hrefValue = link.GetAttributeValue("href", string.Empty);

                // make sure we are inside search results
                if (!hrefValue.ToString().ToUpper().Contains("GOOGLE") && hrefValue.ToString().Contains("/url?q=") && hrefValue.ToString().ToUpper().Contains("HTTPS://"))
                {
                    int index = hrefValue.IndexOf("&");
                    if (index > 0)
                    {
                        hrefValue = hrefValue.Substring(0, index);
                        string possible = hrefValue.Replace("/url?q=", "");
                        bool pos = possible.StartsWith("https://www.walmart.com");
                        if (pos)
                        {
                            links.Add(hrefValue.Replace("/url?q=", ""));
                        }
                    }
                    if (i++ > limit)
                    {
                        break;
                    }
                }
            }
            if (links.Count > 0)
            {
                supplierURL = links[0];
            }
            return supplierURL;
        }


        public static List<string> BingSearch(string search)
        {
            var links = new List<string>();
            try
            {
                IWebDriver driver = new ChromeDriver();
                driver.Navigate().GoToUrl("https://www.bing.com/?toWww=1");
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(60);

                Thread.Sleep(2000);
                driver.FindElement(By.Id("sb_form_q")).SendKeys(search);
                driver.FindElement(By.Id("sb_form_q")).Submit();

                Thread.Sleep(1000);

                foreach (var item in driver.FindElements(By.TagName("a")))
                {
                    var x = item.GetAttribute("href");
                    if (x != null)
                    {
                        links.Add(x);
                    }
                }
                driver.Quit();
            }
            catch (Exception exc)
            {
                string header = "BingSearch";
                string msg = dsutil.DSUtil.ErrMsg(header, exc);
                dsutil.DSUtil.WriteFile(_logfile, msg, "");
            }
            return links;
        }
        public static bool IsDigitsOnly(string str)
        {
            foreach (char c in str)
            {
                if (c < '0' || c > '9')
                    return false;
            }
            return true;
        }
    }
}
