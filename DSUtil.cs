using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

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
            {
                s += i + limiter;
            }
            if (s.Length > 0)
            {
                s = s.Substring(0, s.Length - 1);   // remove last semi colon
            }
            return s;
        }
        public static string ListToDelimited(List<string> list, char limiter)
        {
            string s = "";
            foreach (string i in list)
            {
                s += i + limiter;
            }
            if (s.Length > 0)
            {
                s = s.Substring(0, s.Length - 1);   // remove last semi colon
            }
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
        /// Probably should put this in walmart library - they have funny habit of placing question marks in odd places.
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
                        // BUG here - need to check that (pos + marker.Length <= justText.len)
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
            pos = justText.ToUpper().IndexOf("INTERACTIVE TOUR");
            if (pos > -1)
            {
                help.Add("contains INTERACTIVE TOUR");
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
            pos = justText.ToUpper().IndexOf("VIRTUAL TOUR");
            if (pos > -1)
            {
                help.Add("contains VIRTUAL TOUR");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("CUSTOMER SERVICE");
            if (pos > -1)
            {
                help.Add("contains CUSTOMER SERVICE");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("RECALL");
            if (pos > -1)
            {
                help.Add("contains RECALL");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("SEE OUR");
            if (pos > -1)
            {
                help.Add("contains SEE OUR");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("SHIPPING");
            if (pos > -1)
            {
                help.Add("contains SHIPPING");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("RETURNS");
            if (pos > -1)
            {
                help.Add("contains RETURNS");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("ALSO AVAILABLE");
            if (pos > -1)
            {
                help.Add("contains ALSO AVAILABLE");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("MULTIPLE SIZES");
            if (pos > -1)
            {
                help.Add("contains MULTIPLE SIZES");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("MULTIPLE COLORS");
            if (pos > -1)
            {
                help.Add("contains MULTIPLE COLORS");
                ret = true;
            }

            pos = justText.ToUpper().IndexOf("AVAILABLE COLORS");
            if (pos > -1)
            {
                help.Add("contains AVAILABLE COLORS");
                ret = true;
            }
            pos = justText.ToUpper().IndexOf("AVAILABLE SIZES");
            if (pos > -1)
            {
                help.Add("contains AVAILABLE SIZES");
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
        /// Don't need the actual hyperlink - just does the str contain a hyperlink
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool ContainsHyperlink(string str)
        {
            const string marker = "<a ";
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
        public static bool ContainsScript(string str)
        {
            const string marker = "<script ";
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
        public static List<string> GoogleSearchSelenium(string search)
        {
            var links = new List<string>();
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
                            links.Add(x);
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
            return links; ;
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
        /// <summary>
        /// Look for items most likely want to remove from item description.
        /// </summary>
        /// <param name="strCheck"></param>
        /// <returns></returns>
        public static List<string> GetDescrWarnings(string descriptionHTML)
        {
            var warning = new List<string>();

            // 4.16.2020 not so important for now
            /*
            bool hasOddQuestionMark = dsutil.DSUtil.ContainsQuestionMark(description, out segment);
            if (hasOddQuestionMark)
            {
                warning.Add("Description has odd place question mark -> " + segment);
            }
            */
            bool hasKeyWords = dsutil.DSUtil.ContationsKeyWords(descriptionHTML, out List<string> help);
            if (hasKeyWords)
            {
                foreach (var h in help)
                {
                    warning.Add("Description " + h);
                }
            }
            bool hasDisclaimer = dsutil.DSUtil.ContationsDisclaimer(descriptionHTML);
            if (hasDisclaimer)
            {
                warning.Add("Description contains Disclaimer");
            }
            bool isComputerCamera = IsCameraComputer(descriptionHTML);
            if (isComputerCamera)
            {
                warning.Add("Description computer/camera");
            }
            bool containsEmail = dsutil.DSUtil.StringContainsEmail(descriptionHTML);
            if (containsEmail)
            {
                warning.Add("Description contains email address");
            }
            bool containsHyperlink = dsutil.DSUtil.ContainsHyperlink(descriptionHTML);
            if (containsHyperlink)
            {
                warning.Add("Description contains hyperlink");
            }
            bool containsScript = dsutil.DSUtil.ContainsScript(descriptionHTML);
            if (containsScript)
            {
                warning.Add("Description contains Script");
            }
            return warning;
        }

        /// <summary>
        /// Walmart only offers 14 day returns on computer and cameras, however this does not apply to printers.
        /// We would prefer to start by searching title but presently don't have title so try description.
        /// </summary>
        /// <param name="description"></param>
        /// <returns></returns>
        protected static bool IsCameraComputer(string description)
        {
            bool ret = false;
            string computerMarker = "computer";
            string computerMarker2 = "laptop";
            string cameraMarker = "camera";

            if (string.IsNullOrEmpty(description))
            {
                return false;
            }
            int pos = description.ToUpper().IndexOf(computerMarker.ToUpper());
            if (pos == -1)
            {
                pos = description.ToUpper().IndexOf(computerMarker2.ToUpper());
                if (pos == -1)
                {
                    pos = description.ToUpper().IndexOf(cameraMarker.ToUpper());
                    if (pos > -1)
                    {
                        ret = true;
                    }
                }
                else
                {
                    ret = true;
                }
            }
            else
            {
                ret = true;
            }
            return ret;
        }
        /// <summary>
        /// Get images from supplier from pictureURLS and store them locally at localPath.
        /// </summary>
        /// <param name="supplierPictureURL"></param>
        /// <param name="localPath"></param>
        /// <returns></returns>
        public static List<string> DownloadImages(List<string> supplierPictureURL, string localPath)
        {
            //const string Url = @"http://localhost:51721/productimages/";
            const string Url = @"http://dscruisecontrol.com/scrapeapi/productimages/";

            var localImageURL = new List<string>();
            foreach (var f in supplierPictureURL)
            {
                Uri uri = new Uri(f);
                string filename = System.IO.Path.GetFileName(uri.LocalPath);
                localImageURL.Add(Url + filename);
                string path = localPath + filename;
                using (WebClient webClient = new WebClient())
                {
                    webClient.DownloadFile(f, path);
                }
            }
            return localImageURL;
        }
        public static bool StringContainsEmail(string search)
        {
            Regex emailRegex = new Regex(@"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*",
            RegexOptions.IgnoreCase);

            //find items that matches with our pattern
            MatchCollection emailMatches = emailRegex.Matches(search);

            if (emailMatches.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
