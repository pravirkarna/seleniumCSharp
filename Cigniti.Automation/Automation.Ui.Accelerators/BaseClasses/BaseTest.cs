using System;
using System.Collections.Generic;
using System.Reflection;
using System.Xml;
using Automation.Ui.Accelerators.Enums;
using Automation.Ui.Accelerators.ReportingClassess;
using Automation.Ui.Accelerators.UtilityClasses;
using OpenQA.Selenium;
using OpenQA.Selenium.Remote;


namespace Automation.Ui.Accelerators.BaseClasses
{
    /// <summary>
    /// Description of BaseTest.
    /// </summary>
    public abstract class BaseTest
    {
        public const string BrowserName = "browserName";
        public const string BrowserVersion = "browserVersion";

        protected BaseTest()
        {
        }

        protected BaseTest(RemoteWebDriver driver)
        {
            this.Driver = driver;
        }

        protected BaseTest(XmlNode testNode)
        {
            this.TestCaseNode = testNode;
        }

        protected static BaseTest Instance => Activator.CreateInstance<BaseTest>();


        /// <summary>
        /// Gets or Sets Driver
        /// </summary>
        public RemoteWebDriver Driver { get; set; }

        /// <summary>
        /// Gets or Sets Reporter
        /// </summary>
        public Iteration Reporter { get; set; }

        /// <summary>
        /// Gets or Sets Step
        /// </summary>
        protected string Step
        {
            get
            {
                return Reporter.Chapter.Step.Title;
            }
            set
            {
                Reporter.Add(new Step(value));
            }
        }

        /// <summary>
        /// Gets or Sets Identity of Test Case
        /// </summary>
        public string TestCaseId { get; set; }

        /// <summary>
        /// Gets or Sets Identity of Test Data
        /// </summary>
        public string TestDataId { get; set; }

        /// <summary>
        /// Gets or Sets Test Data as XMLNode
        /// </summary>
        public XmlNode TestDataNode { get; set; }

        /// <summary>
        /// Gets or Sets Test Case as XMLNode
        /// </summary>
        public XmlNode TestCaseNode { get; set; }

        protected T Page<T>() where T : BasePage
        {
            Type pageType = typeof(T);
            return (pageType != null) ? Activator.CreateInstance<T>() : null;
        }

        protected T Page<T>(RemoteWebDriver driver) where T : BasePage, new()
        {
            Type pageType = typeof(T);
            return (T)Activator.CreateInstance(pageType, new object[] { driver });
        }

        protected T Page<T>(XmlNode _testNode) where T : BasePage, new()
        {
            Type pageType = typeof(T);
            return (T)Activator.CreateInstance(pageType, new object[] { _testNode });
        }

        protected T Page<T>(RemoteWebDriver driver, XmlNode _testNode) where T : BasePage, new()
        {
            Type pageType = typeof(T);
            return (T)Activator.CreateInstance(pageType, new object[] { driver, _testNode });
        }

        protected T Page<T>(RemoteWebDriver driver, XmlNode _testNode, Iteration iteration) where T : BasePage, new()
        {
            Type pageType = typeof(T);
            return (T)Activator.CreateInstance(pageType, new object[] { driver, _testNode, iteration });

        }

        protected T Page<T>(RemoteWebDriver driver, XmlNode _testNode, Iteration iteration, string moduleName) where T : BasePage, new()
        {
            Type pageType = typeof(T);
            return (T)Activator.CreateInstance(pageType, new object[] { driver, _testNode, iteration, moduleName });
        }

        public void Execute(TestCase testCaseObject, Dictionary<String, String> browserConfig, XmlNode testDataNode, Iteration iteration, Engine reportEngine)
        {
            try
            {
                Driver = Utility.GetDriver(browserConfig);
                Reporter = iteration;
                TestCaseId = testCaseObject.Title;
                TestDataId = testDataNode.SelectSingleNode(ConfigKey.TestDataId.GetDescription())?.InnerText;
                TestDataNode = testDataNode;
                if (browserConfig[ConfigKey.Target.GetDescription()] == ConfigKey.Local.GetDescription())
                {
                    Reporter.Browser.BrowserName = ((RemoteWebDriver)Driver).Capabilities.GetCapability(BrowserName).ToString();
                    Reporter.Browser.BrowserVersion = ((RemoteWebDriver)Driver).Capabilities.GetCapability(BrowserVersion).ToString();
                }

                if (Reporter.Chapter.Steps.Count == 0)
                    Reporter.Chapters.RemoveAt(0);

                ExecuteTestCase();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                this.Reporter.Chapter.Step.Action.IsSuccess = false;
                this.Reporter.Chapter.Step.Action.TestActExtra(Driver);
            }
            finally
            {
                try
                {
                    this.Reporter.IsCompleted = true;
                    if (this.Reporter.Chapter.Step.Action.Extra == null)
                        this.Reporter.Chapter.Step.Action.Extra = "User defined error <br/> ";
                    this.Reporter.EndTime = TimeZoneInfo.ConvertTime(DateTime.Now, TimeZoneInfo.Local, TimeZoneInfo.Local);

                    lock (reportEngine)
                    {
                        reportEngine.PublishIteration(this.Reporter);
                        reportEngine.Summarize(false);
                    }
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); }

                Driver.Quit();
            }
        }

        /// <summary>	
        /// Executes Test Case, should be overriden by derived
        /// </summary>
        protected virtual void ExecuteTestCase()
        {
            Reporter.Add(new Chapter("Execute Test Case"));
        }

        /// <summary>
        /// Prepares Seed Data, should be overriden by derived
        /// </summary>
        protected virtual void PrepareSeed()
        {
            Reporter.Add(new Chapter("Prepare Seed Data"));
        }
    }
}