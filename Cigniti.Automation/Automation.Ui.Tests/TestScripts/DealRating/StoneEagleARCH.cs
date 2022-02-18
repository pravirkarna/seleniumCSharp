
using Maxim.Automation.Ui.Repository.Constants;
using Maxim.Automation.Ui.Repository.EnumAndConstants;

#region Microsoft references

using System;
#endregion

#region Automation references
using Maxim.Automation.Ui.Accelerators.BaseClasses;
using Maxim.Automation.Ui.Accelerators.ReportingClassess;
using Maxim.Automation.Ui.Accelerators.UtilityClasses;
using Maxim.Automation.Ui.Repository.CommonFunctions;
using Maxim.Automation.Ui.Repository.PageFunctions;
#endregion

namespace Maxim.eTrak.Tests.TestScripts.DealRating
{
    [Script("", "55311", "ETrak", "Deal Rating", "StoneEagle ARCH Rating", "New Vehicle Sale for StoneEagle ARCH dealer", "Regression")]
    class StoneEagleARCH : BaseTest
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

                Step = "Launch 'MenuRC' application";
                var currentPage = commonPage.NavigateToMaximTrakLoginPage();

                Step = "Enter UserName";
                currentPage.EnterUsername(TestDataNode["UserName"].InnerText);

                Step = "Enter Password";
                currentPage.EnterPassword(TestDataNode["Password"].InnerText);

                Step = "Click on Submit button";
                currentPage.ClickOnSubmitButton();

                Step = "Verify Login Success";
                currentPage.VerifyLoginSuccess();

                Step = "Select the dealer from MaximTrak Drop down";
                bool result = currentPage.SelectDealerFromMaximTrakDropDwn(Ratings.StoneEaglearch.GetDescription());

                Step = "Navigate to F&I deals";
                currentPage.ClickFnI();

                Step = "Select New Deal from F&I Page";
                currentPage.NewDeal();

                Step = "Select " + Constants.ArchproductTemplate + " template";
                currentPage.SelectTemplate(Constants.ArchproductTemplate);

                var vinNumber = commonPage.GetVinNumber();

                Step = "Enter the data in the New Deal Wizard -  Deal Worksheet Tab";
                currentPage.KeyInWorkSheetDetails(commonPage.GetCurrentMonthName() + " Stone Eagle", TestDataNode["CustomerLastName"].InnerText,
                                                    TestDataNode["CustomerBirthDate"].InnerText, TestDataNode["SalePrice"].InnerText,
                                                    vinNumber, TestDataNode["MSRP"].InnerText, TestDataNode["Term"].InnerText);

                Step = "Enter the data in New Deal Wizard -  Customer Tab";
                currentPage.KeyInCustomerDetails(TestDataNode["CustomerAddress1"].InnerText, TestDataNode["CustomerAddress2"].InnerText,
                                                TestDataNode["CustomerCity"].InnerText, TestDataNode["CustomerZipCode"].InnerText,
                                                TestDataNode["CustomerHomePhone"].InnerText, TestDataNode["CustomerWorkPhone"].InnerText,
                                                TestDataNode["CustomerEmailAddress"].InnerText);

                Step = "Entering Co Buyer Details - Customer Tab";
                currentPage.KeyInCoBuyerDetails(TestDataNode["CoBuyerFirstName"].InnerText, TestDataNode["CoBuyerLastName"].InnerText, TestDataNode["CoBuyerCountry"].InnerText,
                                             TestDataNode["CoBuyerAddress1"].InnerText, TestDataNode["CoBuyerAddress2"].InnerText, TestDataNode["CoBuyerCity"].InnerText,
                                             TestDataNode["CoBuyerState"].InnerText, TestDataNode["CoBuyerZipCode"].InnerText, TestDataNode["CoBuyerBirthDate"].InnerText,
                                             TestDataNode["CoBuyerHomePhone"].InnerText);

                Step = "Enter the data in New Deal Wizard - Customer Vehicle Info - Vehicle Tab";
                currentPage.KeyInVehicleDetails(TestDataNode["Odometer"].InnerText, TestDataNode["InServiceDate"].InnerText);

                Step = "Enter the data in the New Deal Wizard - Lender Details Tab";    
                currentPage.KeyInLenderDetails(TestDataNode["LenderDetailsName"].InnerText);

                Step = "Click On Save Button";
                currentPage.ClickOnSaveButton();

                Step = "Click on product pricing tab";
                currentPage.ClickOnProductPricingTab();

                Step = "Rate Products";
                currentPage.RateProducts();

                Step = "Select the products and save";
                currentPage.SelectRatedProducts();

                Step = "Click on Electronic Contracting tab";
                currentPage.ClickOnElectronicContractingTab();

                Step = "Complete the Electronic Contracting";
                currentPage.ElectronicContractingCompletion();

                Step = "Select the Registration From Accouting Drop Down";
                currentPage.SelectSubMenuFromAccountingDropDown("Registration");

                Step = "Regester a contract with provider";
                string contractNo = currentPage.RegisterContractWithProvider(vinNumber.Substring(9, 8));

                Step = "Validate the Registration status in All Contracts Tab";
                currentPage.ValidateAllContracts(contractNo, Constants.RegisterdAndApproved);

                Step = "Logout from the application";
                currentPage.LogOutMaximTrak();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}