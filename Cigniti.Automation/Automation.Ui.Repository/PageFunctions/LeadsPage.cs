using Automation.Ui.Accelerators.ReportingClassess;
using Automation.Ui.Repository.CommonFunctions;
using OpenQA.Selenium;

namespace Automation.Ui.Repository.PageFunctions
{
    public partial class SalesForceApplication : Common
    {
        public void CreateNewLead()
        {
            Reporter.Add(new Act("Trying to create New Lead"));
            ClickOnNavObject("Lead");
            ClickOnNewButton("Lead");
            EnterLeadInformation();
            Reporter.Add(VerifyNewLeadCreated() ? new Act("Successfully created New Lead") : new Act("Failed to create New Lead or fields mismatch", false));
        }

        public void CreateNewTaskForLead()
        {
            Reporter.Add(new Act("Trying to create New Task for Lead"));

            EnterNewTaskInformation();
            Reporter.Add(VerifyNewLeadTaskCreated() ? new Act("Successfully created New Lead Task") : new Act("Failed to create New Lead Task", false));            
        }

        public void EnterNewTaskInformation()
        {
            JavaScriptClick(By.XPath("//a[@title='New Task']"), "New Task");            
            SetObjectValue(By.XPath("//label[text()='Subject']/following-sibling::div//input"), "Subject", "Automation Meeting");
            SetObjectValue(By.XPath("//input[@class='inputDate input']"), "Due Date", "12/31/2020");
            JavaScriptClick(By.XPath("//div[@class='bottomBarRight slds-col--bump-left']/button/span[text()='Save']"), "Task Save button");
                    }
        public void EnterLeadInformation()
        {
            Reporter.Add(new Act("Entering New Lead information"));

            //Selection Input
            ObjectClick(By.XPath("//div[@class='uiInput uiInputSelect forceInputPicklist uiInput--default uiInput--select']//span[text()='Lead Status']/parent::span/following-sibling::div//a[text()='New']"), "Lead Status");
            ObjectClick(By.XPath($"//ul[@class='scrollable']//a[text()= '{TestDataNode["LeadStatus"].InnerText}']"), "New");
            ObjectClick(By.XPath("//div[@class='uiInput uiInputSelect forceInputPicklist uiInput--default uiInput--select']//span[text()='Salutation']/parent::span/following-sibling::div//a[text()='--None--']"), "Salutation");
            ObjectClick(By.XPath($"//ul[@class='scrollable']//a[text()= '{TestDataNode["Salutation"].InnerText}']"), "Mr.");
            ObjectClick(By.XPath("//div[@class='uiInput uiInputSelect forceInputPicklist uiInput--default uiInput--select']//span[text()='Rating']/parent::span/following-sibling::div//a[text()='--None--']"), "Rating");
            ObjectClick(By.XPath($"//ul[@class='scrollable']//a[text()= '{TestDataNode["Rating"].InnerText}']"), "Hot");
            ObjectClick(By.XPath("//div[@class='uiInput uiInputSelect forceInputPicklist uiInput--default uiInput--select']//span[text()='Industry']/parent::span/following-sibling::div//a[text()='--None--']"), "Industry");
            ObjectClick(By.XPath($"//ul[@class='scrollable']//a[text()= '{TestDataNode["Industry"].InnerText}']"), "Banking");
            ObjectClick(By.XPath("//div[@class='uiInput uiInputSelect forceInputPicklist uiInput--default uiInput--select']//span[text()='Lead Source']/parent::span/following-sibling::div//a[text()='--None--']"), "Lead Source");
            ObjectClick(By.XPath($"//ul[@class='scrollable']//a[text()= '{TestDataNode["LeadSource"].InnerText}']"), "Advertisement");

            //Textbox input
            SetObjectValue(By.XPath("//span[text()='First Name']/parent::label/following-sibling::input"), "FirstName", TestDataNode["LeadFirstName"].InnerText);
            SetObjectValue(By.XPath("//span[text()='Last Name']/parent::label/following-sibling::input"), "LastName", TestDataNode["LeadLastName"].InnerText);            
            SetObjectValue(By.XPath("//span[text()='Title']/parent::label/following-sibling::input"), "Title", TestDataNode["Title"].InnerText);
            SetObjectValue(By.XPath("//span[text()='Email']/parent::label/following-sibling::input"), "Email", TestDataNode["Email"].InnerText);
            SetObjectValue(By.XPath("//span[text()='Phone']/parent::label/following-sibling::input"), "Phone", TestDataNode["Phone"].InnerText);
            SetObjectValue(By.XPath("//span[text()='Mobile']/parent::label/following-sibling::input"), "Mobile", TestDataNode["Mobile"].InnerText);
            SetObjectValue(By.XPath("//span[text()='Company']/parent::label/following-sibling::input"), "Company", TestDataNode["Company"].InnerText);
            SetObjectValue(By.XPath("//span[text()='No. of Employees']/parent::label/following-sibling::input"), "No. of Employees", TestDataNode["NoofEmployees"].InnerText);
            SetObjectValue(By.XPath("//span[text()='Website']/parent::label/following-sibling::input"), "Website", TestDataNode["Website"].InnerText); 

            ObjectClick(By.XPath("//span[text()='Search Address']"), "Search Address");
            Wait(1000);
           // SetObjectValue(By.XPath("//input[@placeholder='Enter address']"), "Enter Address", "Madhapur");
            Wait(1000);
            SearchAddress(TestDataNode["Address"].InnerText);
            ObjectClick(By.XPath("//ul[@class='lookup__list  visible']//a[1]"), "First match of the address");
            Wait(3000);
            ObjectClick(By.XPath("//div[@class='actionsContainer']//button//span[text()='Save']"));
            WaitForPageLoad();
        }
        public void SearchAddress(string address="")
        {
            var element = WaitForElementVisible((By.XPath("//input[@placeholder='Enter address']")));
            for(int i=0;i<(address.Length);)
            {
                element.SendKeys(string.Concat(address[i],address[i+1]));
                i = i + 2;
                Wait(1000);
            }
        }
        public void ClickOnNewButton(string navObjectName="Lead")
        {
            Reporter.Add(new Act($"Clicking on New {navObjectName} button"));

            ObjectClick(By.XPath("//a/div[text()='New']"), navObjectName);
            WaitForPageLoad();
        }

        public bool VerifyNewLeadTaskCreated()
        {
            Reporter.Add(new Act("Verifying if New Lead created"));

           return ValidateIfExists(By.XPath($"//div[@class='timelineRow slds-media__body slds-grid']//a[text()='{TestDataNode["LeadTaskName"].InnerText}']"), "New Lead Task");
        }

        public bool VerifyNewLeadCreated()
        {
            Reporter.Add(new Act("Verifying if New Lead created"));

            if (ValidateIfExists(By.XPath($"//lightning-formatted-name[text()='{TestDataNode["LeadFullName"].InnerText}']"), "Lead Name")
               && ValidateIfExists(By.XPath($"//force-highlights-details-item//lightning-formatted-text[text()='{TestDataNode["Company"].InnerText}']"), "Company Name")
               && ValidateIfExists(By.XPath($"//force-highlights-details-item//lightning-formatted-text[text()='{TestDataNode["Title"].InnerText}']"), "Title")
               && ValidateIfExists(By.XPath($"//force-highlights-details-item/div//p//lightning-formatted-phone/a[@href='tel:{TestDataNode["Phone"].InnerText}']"), "Phone number")
               && ValidateIfExists(By.XPath($"//emailui-formatted-email-lead//a[text()='{TestDataNode["Email"].InnerText}']"), "Email")
               && ValidateIfExists(By.XPath($"//a[text()='{TestDataNode["Website"].InnerText}']"), "Company Site")
               && ValidateIfExists(By.XPath($"//span[text()='Industry']/parent::div/following-sibling::div//span[text()='{TestDataNode["Industry"].InnerText}']"), "Industry")
               && ValidateIfExists(By.XPath($"//span[text()='No. of Employees']/parent::div/following-sibling::div//span[text()='{TestDataNode["NoofEmployees"].InnerText}']"), "No.of Employees"))
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
