#region Microsoft Reference
using System;
using System.Xml;
using System.Configuration;
using System.IO;
using System.Net;
using System.Threading;
#endregion

#region Selenium Reference
using OpenQA.Selenium.Remote;
#endregion

#region Cigniti Automation Reference
using System.Collections.Generic;
using OpenQA.Selenium;
using static Automation.Ui.Repository.Constants.Constants;
using Automation.Ui.Accelerators.ReportingClassess;
using Automation.Ui.Accelerators.UtilityClasses;
using Automation.Ui.Accelerators.BaseClasses;
using System.Reflection;
using Automation.Ui.Repository.PageFunctions;
using Automation.Ui.Repository.Enums;

namespace Automation.Ui.Repository.CommonFunctions
{

    public class Common : BasePage
    {
        public string Name { get; set; }

        public static List<string> VinList { get; set; } = new List<string>();

        public static object Locker { get; set; } = new object();


        public Common()
        {
        }

        public Common(RemoteWebDriver driver)
        {
            this.Driver = driver;
        }

        public Common(XmlNode testNode)
        {
            this.TestDataNode = testNode;
        }

        public Common(RemoteWebDriver driver, XmlNode testNode)
        {
            this.Driver = driver;
            this.TestDataNode = testNode;
        }

        public Common(RemoteWebDriver driver, XmlNode testNode, Iteration iteration)
        {
            this.Driver = driver;
            this.TestDataNode = testNode;
            this.Reporter = iteration;

            this.PageObjects = Locator.LoadPageObjects(Locator.ModuleName);
        }

        public Common(RemoteWebDriver driver, XmlNode testNode, Iteration iteration, string moduleName)
        {
            this.Driver = driver;
            this.TestDataNode = testNode;
            this.Reporter = iteration;
            this.PageObjects = Locator.LoadPageObjects(moduleName);
            Locator.ModuleName = moduleName;
        }

        public SalesForceApplication NavigateToSalesForceLoginPage()
        {
            var applicationUrl = ConfigurationManager.AppSettings.Get("URL");
            Reporter.Add(new Act("Navigating to " + applicationUrl));
            NavigateToUrl(applicationUrl);
            WaitForPageLoad();

            return new SalesForceApplication(Driver, TestDataNode, Reporter);
        }

        public void VerifyPageLoad(int maxTime = 120)
        {
            Reporter.Add(new Act(String.Format("Wait for the current page to load fully")));
            var count = 0;
            for (count = 0; count < maxTime; count++)
            {
                if (IsWebElementDisplayed(Locator.GetLocator(Page.Generic.GetDescription(), HomePageObjects.HomeNavBar.GetDescription()), 1))
                {
                    Reporter.Add(new Act("Page is loaded successfully"));
                    break;
                }

                Wait(1000);
            }

            if (count >= maxTime)
                throw new Exception(($"{MethodBase.GetCurrentMethod()}: Failed to load page"));
        }

        public string GetXpathString(string locatorName, string elementName)
        {
            var xpath = string.Empty;
            var locator = Locator.GetLocator(locatorName, elementName).ToString().Split(':');

            Array.ForEach(locator, l =>
            {
                if (!l.Equals("By.XPath"))
                    xpath += string.IsNullOrEmpty(l) ? "::" : l;
            });

            return xpath;
        }

        public void LogOutMaximTrak()
        {
            Reporter.Add(new Act(string.Format("Trying to Logout of the MaximTrak Application")));

            //ObjectClick(Locator.GetLocator(PAGE.GENERIC.GetDescription(), GENERICOBJECTS.LOGOUT.GetDescription()),
            //                                GENERICOBJECTS.LOGOUT.GetDescription(), 5);
            //WaitForPageLoad();
            //Wait(3000);
            //VerifyPageLoad();

            //if (ValidateIfExists(
            //        Locator.GetLocator(PAGE.LOGIN.GetDescription(), LOGINOBJECTS.USERNAME.GetDescription()),
            //        LOGINOBJECTS.USERNAME.GetDescription())
            //    &&
            //    ValidateIfExists(Locator.GetLocator(PAGE.LOGIN.GetDescription(), LOGINOBJECTS.PASSWORD.GetDescription()),
            //    LOGINOBJECTS.PASSWORD.GetDescription()))
            //{
            //    Reporter.Add(new Act($"User: {TestDataNode[Constants.Constants.Name]?.InnerText} Successfully Logged out"));
            //}
            //else
            //{
            //    throw new Exception($"User: {TestDataNode[Constants.Constants.Name]?.InnerText} failed to log-out");
            //}
        }

        /// <summary>
        /// Gets Month name
        /// </summary>
        /// <returns>Returns first three characters of the current month.</returns>
        public string GetCurrentMonthName() => DateTime.Now.ToString("MMMM").Substring(0, 3);


        /// <summary>
        /// Finds the xpath with visible index
        /// </summary>
        /// <param name="locator"></param>
        /// <returns>Xpath with Index</returns>
        public By GetXPathWithVisibleElementIndex(string locator)
        {
            try
            {
                var elements = Driver.FindElements(By.XPath(locator));
                int elementIndex = 1;
                string xpath = string.Empty;

                foreach (var item in elements)
                {
                    xpath = $"({locator})[{elementIndex}]";
                    if (IsWebElementDisplayed(By.XPath(xpath)))
                    {
                        break;
                    }

                    elementIndex++;
                }

                return By.XPath(xpath);
            }
            catch (Exception ex)
            {
                throw new Exception(
                    $"Failed to find element index at GetXPathWithVisibleElementIndex() {ex.Message}");
            }
        }

        #endregion

        #region Endpoints Helper methods

        private string GetSoapResponse(string fileName, out string dealerNumber)
        {
            Reporter.Add(new Act($"SOAP Response requested at {DateTime.Now:HH:mm:ss}"));

            var path = $"{Directory.GetCurrentDirectory()}\\{ fileName }";

            var xmlContent = File.ReadAllText(path);
            dealerNumber = GenerateRandomNumber(3);
            var tempxmlContent = string.Format(xmlContent, dealerNumber, dealerNumber);

            using (var writer = new StreamWriter(path))
            {
                writer.Write(tempxmlContent);
            }
            Wait(2000);
            var webRequest = CreateWebRequest();
            var soapReqBody = new XmlDocument();
            soapReqBody.Load($"{Directory.GetCurrentDirectory()}\\{fileName}");
            Wait(2000);
            using (var requestStream = webRequest.GetRequestStream())
            {
                soapReqBody.Save(requestStream);
                Wait(1000);
            }

            string result;
            try
            {
                using (var response = webRequest.GetResponse())
                {
                    Wait(3000);
                    using (var rd = new StreamReader(response.GetResponseStream()))
                    {
                        result = rd.ReadToEnd();
                        result += "Status: OK";
                        Wait(2000);
                    }
                }
            }
            catch (WebException e)
            {
                var message = new StreamReader(e.Response.GetResponseStream()).ReadToEnd();
                Console.WriteLine(message);
                throw;
            }

            return result;
        }

        private HttpWebRequest CreateWebRequest()
        {
            var webRequest = (HttpWebRequest)WebRequest.Create(TestDataNode[Constants.Constants.EndPointDetails]?.SelectSingleNode("EndPointURL")?.InnerText ?? string.Empty);
            webRequest.Headers.Add(TestDataNode[Constants.Constants.EndPointDetails]?.SelectSingleNode("SoapAction")?.InnerText ?? string.Empty);
            webRequest.ContentType = TestDataNode[Constants.Constants.EndPointDetails]?.SelectSingleNode("ContentType")?.InnerText;
            webRequest.Accept = "*/*";
            webRequest.Method = TestDataNode[Constants.Constants.EndPointDetails]?.SelectSingleNode("Method")?.InnerText ?? string.Empty;
            webRequest.KeepAlive = true;
            webRequest.Timeout = Timeout.Infinite;
            return webRequest;
        }

        #endregion
    }
}