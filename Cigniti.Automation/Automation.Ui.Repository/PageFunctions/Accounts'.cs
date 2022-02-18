using Automation.Ui.Accelerators.ReportingClassess;
using Automation.Ui.Repository.CommonFunctions;
using OpenQA.Selenium;

namespace Automation.Ui.Repository.PageFunctions
{
    public partial class SalesForceApplication : Common
    {
        public void CreateNewAccount()
        {
            Reporter.Add(new Act("Trying to create new account"));

            ClickOnNavObject("Account");
            ClickOnNewButton("Account");
            EnterAccountInformtion();
            Reporter.Add(VerifyNewAccountCreated() ? new Act("Successfully created New Account") : new Act("Failed to create New Account or fields mismatch", false));
        }

        public void CreateNewOpportunities()
        {
            Reporter.Add(new Act("Trying to create New Opportunities for Account"));

            EnterAccountOpportunitiesInformation();
            Reporter.Add(VerifyNewOpportunity() ? new Act("Successfully created New Account Opportunity") : new Act("Failed to create New Account Opportunity", false));
        }

        public void EnterAccountInformtion()
        {
            Reporter.Add(new Act("Entering New Account information"));

            // SetObjectValue(By.XPath("//span[text()='Account Name']/parent::label/following-sibling::div//input"), "Account Name","Cigniti");
            var accoutName = WaitForElementVisible(By.XPath("//span[text()='Account Name']/parent::label/following-sibling::div//input"));
            accoutName.SendKeys("Cig");
            Wait(1000);
            accoutName.SendKeys("ni");
            Wait(1000);
            accoutName.SendKeys("ti");
            Wait(2000);
            ObjectClick(By.XPath("//ul[@class='lookup__list  visible']//a[1]"), "Select from account list");
            ObjectClick(By.XPath("//div[@class='uiInput uiInputSelect forceInputPicklist uiInput--default uiInput--select']//span[text()='Type']/parent::span/following-sibling::div//a[text()='--None--']"), "Type");
            ObjectClick(By.XPath($"//ul[@class='scrollable']//a[text()= '{TestDataNode["AccountType"].InnerText}']"), "Select Type from list");
            ObjectClick(By.XPath("//div[@class='uiInput uiInputSelect forceInputPicklist uiInput--default uiInput--select']//span[text()='Industry']/parent::span/following-sibling::div//a[text()='--None--']"), "Industry");
            ObjectClick(By.XPath($"//ul[@class='scrollable']//a[text()= '{TestDataNode["Industry"].InnerText}']"), "Select Type from list");

            ObjectClick(By.XPath("//span[text()='Billing Address']/parent::legend/following-sibling::div//span[text()='Search Address']"), "Billing Address");
            Wait(1000);            
            SearchAddress(TestDataNode["Address"].InnerText);
            Wait(1000);
            ObjectClick(By.XPath("//ul[@class='lookup__list  visible']//a[1]"), "First match of the address");
            Wait(3000);
            ObjectClick(By.XPath("//span[text()='Shipping Address']/parent::legend/following-sibling::div//span[text()='Search Address']"), "Shipping Address");
            Wait(1000);
            SearchAddress(TestDataNode["Address"].InnerText);
            ObjectClick(By.XPath("//ul[@class='lookup__list  visible']//a[1]"), "First match of the address");
            Wait(3000);

            ObjectClick(By.XPath("//div[@class='button-container-inner slds-float_right']/button/span[text()='Save']"), "Save Account");
            WaitForPageLoad();
        }

        public void EnterAccountOpportunitiesInformation()
        {
            Reporter.Add(new Act("Entering New Opportunities information for account"));
            Wait(2000);
            JavaScriptClick(By.XPath("//span[text()='Opportunities']/../../../parent::header/following-sibling::div//a[@title='New']"), "New Opportunity");
            SetObjectValue(By.XPath("//span[text()='Opportunity Name']/parent::label/following-sibling::input"), "Opportunity Name", TestDataNode["OpportunityName"].InnerText);
            SetObjectValue(By.XPath("//span[text()='Close Date']/parent::label/following-sibling::div/input"),"Close Date", TestDataNode["CloseDate"].InnerText);
            ObjectClick(By.XPath("//span[text()='Type']/parent::span/following-sibling::div//a[text()='--None--']"), "Type selection");
            ObjectClick(By.XPath($"//ul[@class='scrollable']//a[text()= '{TestDataNode["TypeSelection"].InnerText}']"), "Type option");
            ObjectClick(By.XPath("//span[text()='Stage']/parent::span/following-sibling::div//a[text()='--None--']"), "Stage");
            ObjectClick(By.XPath($"//ul[@class='scrollable']//a[text()= '{TestDataNode["Stage"].InnerText}']"), "Stage option");

            ObjectClick(By.XPath("//div[@class='button-container-inner slds-float_right']/button/span[text()='Save']"), "Save Opportunity");
            WaitForPageLoad();
        }
        public bool VerifyNewAccountCreated()
        {
            Reporter.Add(new Act("Verifying if New Account created"));
            if (ValidateIfExists(By.XPath("//a[text()='https://www.cigniti.com/']"), "Company Site") &&
               ValidateIfExists(By.XPath($"//force-highlights-details-item/div//p//lightning-formatted-phone/a[@href='tel:{TestDataNode["AccountPhone"].InnerText}']"), "Phone Number") &&
               ValidateIfExists(By.XPath("//div/following-sibling::slot//span[text()='Cigniti Inc']"), "Account Name"))
            {
                return true;
            }
            return false;
        }
        public bool VerifyNewOpportunity()
        {
            Reporter.Add(new Act("Verifying if New Lead created"));

            return ValidateIfExists(By.XPath($"//div[@class='outputLookupContainer forceOutputLookupWithPreview']/a[text()='{TestDataNode["OpportunityName"].InnerText}']"), "New Opportunity");
        }
    }
}
