using Automation.Ui.Accelerators.ReportingClassess;
using Automation.Ui.Accelerators.UtilityClasses;
using Automation.Ui.Repository.CommonFunctions;
using Automation.Ui.Repository.Enums;
using OpenQA.Selenium.Remote;
using System;
using System.Reflection;
using System.Xml;

namespace Automation.Ui.Repository.PageFunctions
{
    /// <summary>
    ///  Represents PageFunction. Inherates from BasePage.
    /// </summary>
    public partial class SalesForceApplication : Common
    {
        /// <summary>
        /// Constructor without parameters
        /// </summary>
        public SalesForceApplication()
        {
        }

        public SalesForceApplication(RemoteWebDriver driver)
        {
            this.Driver = driver;
        }

        public SalesForceApplication(XmlNode testNode)
        {
            this.TestDataNode = testNode;
        }

        public SalesForceApplication(RemoteWebDriver driver, XmlNode testNode)
        {
            this.Driver = driver;
            this.TestDataNode = testNode;
        }

        public SalesForceApplication(RemoteWebDriver driver, XmlNode testNode, Iteration iteration)
        {
            this.Driver = driver;
            this.TestDataNode = testNode;
            this.Reporter = iteration;
        }

        public SalesForceApplication(RemoteWebDriver driver, XmlNode testNode, Iteration iteration, string moduleName)
            : base(driver, testNode, iteration, moduleName)
        {
            this.Driver = driver;
            this.TestDataNode = testNode;
            this.Reporter = iteration;
        }

        /// <summary>
        /// Enters UserName in username textbox.
        /// </summary>
        /// <param name="userName">UserName</param>
        public void EnterUsername(string userName)
        {
            try
            {
                Reporter.Add(new Act($"Entering UserName: {userName} in textbox"));

                SetObjectValue(Locator.GetLocator(Page.Generic.GetDescription(), LoginObjects.UserName.GetDescription()),
                               LoginObjects.UserName.GetDescription(),
                               userName,
                               5);

            }
            catch (Exception ex)
            {
                Console.WriteLine(($"Failed at {MethodBase.GetCurrentMethod()} {ex.Message}\n{ex.StackTrace}"));
            }
        }

        /// <summary>
        /// Enters Password in Password textbox.
        /// </summary>
        /// <param name="password">UserName password</param>
        public void EnterPassword(string password)
        {
            try
            {
                Reporter.Add(new Act($"Entering Password: {password} in textbox"));

                SetObjectValue(Locator.GetLocator(Page.Generic.GetDescription(), LoginObjects.Password.GetDescription()),
                              LoginObjects.UserName.GetDescription(),
                              password,
                              5);
            }
            catch (Exception ex)
            {
                Console.WriteLine(($"Failed at {MethodBase.GetCurrentMethod()} {ex.Message}\n{ex.StackTrace}"));
            }
        }

        /// <summary>
        /// Click On Submit Button
        /// </summary>
        public void ClickOnSubmitButton()
        {
            try
            {
                Reporter.Add(new Act($"Trying to click on submit button"));

                ObjectClick(Locator.GetLocator(Page.Generic.GetDescription(), LoginObjects.SubmitBtn.GetDescription()), "Submit Button");

                Wait(5000);
            }
            catch (Exception ex)
            {
                Console.WriteLine(($"Failed at {MethodBase.GetCurrentMethod()} {ex.Message}\n{ex.StackTrace}"));
            }
        }

        /// <summary>
        /// Verifies if Login is successful
        /// </summary>
        public void VerifyLoginSuccess()
        {
            try
            {
                Reporter.Add(new Act($"Verifying if login is successful"));

                WaitForPageLoad();
                VerifyPageLoad();
                if (ValidateIfExists(Locator.GetLocator(Page.Generic.GetDescription(), HomePageObjects.HomeNavBar.GetDescription())
                    , "Nav bar"))
                {

                    Reporter.Add(new Act($"Successfully Logged in"));
                }
                else
                {
                    Reporter.Add(new Act($"Failed to log-in", false, Driver));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(($"Failed at {MethodBase.GetCurrentMethod()} {ex.Message}\n{ex.StackTrace}"));
            }
        }

        public void Login()
        {
            EnterUsername(TestDataNode["UserName"].InnerText);
            EnterPassword(TestDataNode["Password"].InnerText);
            ClickOnSubmitButton();
            VerifyLoginSuccess();
        }

        public void LogOut()
        {
            Reporter.Add(new Act($"Trying to logout."));
            ClickonUserIcon();
            JavaScriptClick(Locator.GetLocator(Page.Generic.GetDescription(), HomePageObjects.LogOut.GetDescription()), "User Icon");
            WaitForPageLoad();
            Wait(1000);
            if (ValidateIfExists(Locator.GetLocator(Page.Generic.GetDescription(), LoginObjects.UserName.GetDescription()), "UserName Textbox")
                &&
                ValidateIfExists(Locator.GetLocator(Page.Generic.GetDescription(), LoginObjects.Password.GetDescription()), "Password Textbox"))
            {

                Reporter.Add(new Act($"Successfully Logged out"));
            }
            else
            {
                Reporter.Add(new Act($"Failed to Log out", false, Driver));
            }
        }

        public void ClickonUserIcon()
        {
            Reporter.Add(new Act($"Clicking on User icon."));
            Wait(2000);
            JavaScriptClick(Locator.GetLocator(Page.Generic.GetDescription(), HomePageObjects.UserIcon.GetDescription()), "User Icon");
            Wait(2000);
        }
    }
}