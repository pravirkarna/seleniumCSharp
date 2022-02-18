using System;
using Automation.Ui.Accelerators.BaseClasses;
using Automation.Ui.Accelerators.ReportingClassess;
using Automation.Ui.Accelerators.UtilityClasses;
using Automation.Ui.Repository.CommonFunctions;

namespace Automation.Ui.Tests.TestScripts.DemoScripts
{
    [Script("", "SalesForce", "Login", "LoginToSalesForce", "Login to SalesForce account with user credentials", "Smoke")]
    class Login : BaseTest
    {
        /// <summary>
        ///  overriden Execute TestCase
        /// </summary>
        protected override void ExecuteTestCase()
        {
            try
            {
                Reporter.Add(new Chapter($"Execute test case- '{this.GetType().Name}'"));

                var commonPage = Page<Common>(Driver, TestDataNode, Reporter);

                Step = "Launch the application";
                var currentPage = commonPage.NavigateToSalesForceLoginPage();

                Step = "Login to SalesForce account";
                currentPage.Login();

                Step = "Logout from SalesForce account";
                currentPage.LogOut();
            }
            catch (Exception ex)
            {
                Reporter.Add(new Act($"Stack: {ex.StackTrace} Message: {ex.Message}", false));
            }
        }
    }
}