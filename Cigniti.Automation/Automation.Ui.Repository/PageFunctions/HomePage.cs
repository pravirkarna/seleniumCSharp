using Automation.Ui.Accelerators.ReportingClassess;
using Automation.Ui.Repository.CommonFunctions;
using OpenQA.Selenium;

namespace Automation.Ui.Repository.PageFunctions
{
    /// <summary>
    ///  Represents PageFunction. Inherates from BasePage.
    /// </summary>
    public partial class SalesForceApplication : Common
    {

        /// <summary>
        /// Click Home Button.
        /// </summary>
        public void ClickHomeIcon()
        {
           
        }

        public void ClickOnNavObject(string navObjectName)
        {
            Reporter.Add(new Act($"Clicking on {navObjectName}"));

            ObjectClick(By.XPath($"//one-app-nav-bar-item-root[@data-id='{navObjectName}']"), navObjectName);
            WaitForPageLoad();
        }
        public void ClickOnNavObjectCarrot(string navObjectName)
        {
            Reporter.Add(new Act($"Clicking on {navObjectName} carrot"));

            ObjectClick(By.XPath($"//nav[@class='slds-context-bar__secondary navCenter']//a[@href='/lightning/o/Lead/home']/following-sibling::one-app-nav-bar-item-dropdown//a[@role='button']"), $"{ navObjectName} carrot");
            WaitForPageLoad();
        }
    }
}