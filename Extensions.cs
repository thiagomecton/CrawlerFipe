using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CrawlerFipe
{
    public static class Extensions
    {
        public static bool GetElement(this RemoteWebDriver driver, By by, out IWebElement element, int tentativas = 1)
        {
            int tentativa = 1;
            element = null;

            while (tentativa <= tentativas)
            {
                try
                {
                    WebDriverWait waitFor = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
                    Func<IWebDriver, IWebElement> waitForElement = new Func<IWebDriver, IWebElement>((IWebDriver Web) =>
                    {
                        return driver.FindElement(by);
                    });

                    element = waitFor.Until(waitForElement);
                    return true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                tentativa++;

                if (tentativa <= tentativas)
                {
                    Thread.Sleep(2000);
                }
            }

            return false;
        }

        public static byte[] ToBytes(this Stream input)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                input.CopyTo(ms);
                return ms.ToArray();
            }
        }
    }
}
