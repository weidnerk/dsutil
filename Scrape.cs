using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web;

namespace dsutil
{
    public static class Scrape
    {
        readonly static string _logfile = "log.txt";

        public static string NavigateToTransHistory(string URL)
        {
            string output = null;
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("headless");
            IWebDriver driver = new ChromeDriver();
            try
            {
                Thread.Sleep(2000);
                driver.Navigate().GoToUrl(URL);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(60);
                Thread.Sleep(2000);

                output = driver.FindElement(By.TagName("html")).GetAttribute("innerHTML");
                output = driver.PageSource;

                driver.Quit();
            }
            catch (Exception exc)
            {
                driver.Quit();
                string msg = dsutil.DSUtil.ErrMsg("NavigateToTransHistory", exc);
                dsutil.DSUtil.WriteFile(_logfile, msg, "");
            }
            return output;
        }

    }
}