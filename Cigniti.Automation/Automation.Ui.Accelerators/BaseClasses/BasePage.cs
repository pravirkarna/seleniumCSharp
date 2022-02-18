using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Xml;
using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System.Data;
using System.Text;
using Automation.Ui.Accelerators.UtilityClasses;
using Automation.Ui.Accelerators.ReportingClassess;
using System.Reflection;

namespace Automation.Ui.Accelerators.BaseClasses
{
    /// <summary>
    /// This is the Super class for all pages
    /// </summary>
    /// 
    public abstract class BasePage
    {
        /// <summary>
        /// Get the browser configuration details
        /// </summary>
        public Dictionary<string, string> BrowserConfig = Utility.BrowserConfig;

        /// <summary>
        /// Gets or Sets Driver
        /// </summary>
        public RemoteWebDriver Driver { get; set; }

        /// <summary>
        /// Gets or Sets Test Data as XMLNode
        /// </summary>
        public XmlNode TestDataNode { get; set; }

        /// <summary>
        /// Gets or Sets Reporter
        /// </summary>
        public Iteration Reporter { get; set; }

        public Dictionary<string, Dictionary<string, Dictionary<string, List<BrowserFinder>>>> PageObjects { get; set; }

        public const string SCROLLINTOVIEW = @"(arguments[0].scrollIntoView(true));";

        public static readonly string HighliteBorderScript = $@"arguments[0].style.cssText = 'border-width: 4px; border-style: solid; border-color: {"orange"}';";

        public BasePage()
        {

        }

        public BasePage(RemoteWebDriver driver)
        {
            Driver = driver;
        }

        public BasePage(XmlNode testNode)
        {
            TestDataNode = testNode;
        }

        public BasePage(RemoteWebDriver driver, XmlNode testNode)
        {
            Driver = driver;
            TestDataNode = testNode;
        }

        public BasePage(RemoteWebDriver driver, XmlNode testNode, Iteration iteration, string moduleName)
        {
            Driver = driver;
            TestDataNode = testNode;
            Reporter = iteration;
        }


        internal IWebElement GetNativeElement(By lookupBy, int maxWaitTime = 60)
        {
            IWebElement element = null;

            for (int i = 0; i < maxWaitTime; i++)
            {
                try
                {
                    element = Driver.FindElement(lookupBy);
                    if (element == null) continue;
                    IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)Driver;
                    jsExecutor.ExecuteScript(HighliteBorderScript, new object[] { element });
                    jsExecutor.ExecuteScript(string.Format(SCROLLINTOVIEW), new object[] { element });
                    break;
                }
                catch
                {
                    Thread.Sleep(1000);
                }
            }

            return element;
        }

        internal IList<IWebElement> GetNativeElements(By lookupBy, int maxWaitTime = 60)
        {
            IList<IWebElement> element = null;

            for (int i = 0; i < maxWaitTime; i++)
            {
                try
                {
                    element = Driver.FindElements(lookupBy);
                    if (element.Count > 0) break;
                }
                catch
                {
                    Thread.Sleep(1000);
                }
            }

            return element;
        }

        internal IWebElement GetNativeElementInElement(IWebElement parentElement, By lookupBy, int maxWaitTime = 60)
        {
            try
            {
                var element = parentElement.FindElement(lookupBy);
                if (element != null)
                {

                    var jsExecutor = (IJavaScriptExecutor)Driver;
                    jsExecutor.ExecuteScript(HighliteBorderScript, new object[] { element });
                    jsExecutor.ExecuteScript(string.Format(SCROLLINTOVIEW), new object[] { element });
                    return element;
                }
            }
            catch
            {
                Thread.Sleep(1000);
            }

            return null;
        }


        public IWebElement WaitForElementVisible(By lookupBy, int maxWaitTime = 60)
        {
            try
            {
                var element = new WebDriverWait(Driver, TimeSpan.FromSeconds(maxWaitTime)).Until(condition =>
                {
                    try
                    {
                        return Driver.FindElement(lookupBy);
                    }
                    catch (Exception)
                    {
                        // This try/catch block is required to continue the execution though the element is not found.
                        return null;
                    }
                });

                if (element == null) return element;
                var jsExecutor = (IJavaScriptExecutor)Driver;
                jsExecutor.ExecuteScript(HighliteBorderScript, new object[] { element });
                jsExecutor.ExecuteScript(string.Format(SCROLLINTOVIEW), new object[] { element });

                return element;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return null;
            }
        }

        /// <summary>
        /// Verifies weather the element is present on the UI or not.
        /// </summary>
        /// <param name="lookupBy"></param>
        /// <param name="maxWaitTime"></param>
        /// <returns>TRUE if element displayed</returns>
        public bool IsWebElementDisplayed(By lookupBy, int maxWaitTime = 30)
        {
            try
            {
                Wait(maxWaitTime);
                return Driver.FindElement(lookupBy)?.Displayed ?? false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error Message: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Gets the Element value
        /// </summary>
        /// <param name="lookupBy"></param>
        /// <param name="attributeName"></param>
        /// <param name="maxWaitTime"></param>
        /// <returns>string of the element</returns>
        public string GetElementValue(By lookupBy, string attributeName = "", int maxWaitTime = 30)
        {
            try
            {
                Wait(maxWaitTime);
                return attributeName == string.Empty ? Driver.FindElement(lookupBy).Text : Driver.FindElement(lookupBy).GetAttribute(attributeName);
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        internal IWebElement WaitForElementVisibleWithoutHighLight(By lookupBy, int maxWaitTime = 60)
        {
            var element = new WebDriverWait(Driver, TimeSpan.FromSeconds(maxWaitTime)).Until(ExpectedConditions.ElementIsVisible(lookupBy));
            if (element == null) return null;
            var jsExecutor = (IJavaScriptExecutor)Driver;
            jsExecutor.ExecuteScript(HighliteBorderScript, new object[] { element });
            jsExecutor.ExecuteScript(string.Format(@"$(arguments[0].scrollIntoView(true));"), new object[] { element });

            return element;
        }


        public void ElementFocus(By lookupBy, int maxWaitTime = 60)
        {
            IWebElement element = WaitForElementVisible(lookupBy, maxWaitTime);
            new Actions(Driver).MoveToElement(element).Perform();
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)Driver;
            jsExecutor.ExecuteScript($"javascript:window.scrollBy({0},{(element.Location.Y - 200)})");
        }

        internal IWebElement WaitForElementVisible(IWebElement WebElement, int maxWaitTime = 60)
        {
            var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(maxWaitTime));
            wait.Until(webEle => WebElement);
            if (WebElement == null) return WebElement;
            var jsExecutor = (IJavaScriptExecutor)Driver;
            jsExecutor.ExecuteScript(HighliteBorderScript, new object[] { WebElement });
            jsExecutor.ExecuteScript(@"$(arguments[0].scrollIntoView(true));", new object[] { WebElement });

            return WebElement;
        }

        public void WaitForElementNotDisplayed(By by)
        {
            for (var i = 0; i < 60; i++)
            {
                if (Driver.FindElement(by)?.Displayed ?? true)
                    Thread.Sleep(1000);
                else
                    break;
            }
        }

        internal IList<IWebElement> GetNativeElementsInElement(IWebElement parentElement, By lookupBy, int maxWaitTime = 60)
        {
            IList<IWebElement> element = null;

            for (var i = 0; i < maxWaitTime; i++)
            {
                try
                {
                    element = parentElement.FindElements(lookupBy);
                    if (element.Count > 0) break;
                }
                catch
                {
                    Thread.Sleep(1000);
                }
            }

            return element;
        }

        public void ObjectClick(By lookupBy, int maxWaitTime = 60)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);

            if (element == null) return;
            Thread.Sleep(2000);
            element.Click();
        }

        public void SetValueToObject(By lookupBy, string strInputValue, int maxWaitTime = 60)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            if (element == null) return;
            element.Clear();
            element.SendKeys(strInputValue);

        }

        public void SetValueToObjectWithoutHighLight(By lookupBy, string strInputValue, int maxWaitTime = 60)
        {
            var element = WaitForElementVisibleWithoutHighLight(lookupBy, maxWaitTime);
            if (element == null) return;
            element.Clear();
            element.SendKeys(strInputValue);
        }

        public void ObjectSelectValue(By lookupBy, string strInputValue, string selectBy = "text", int maxWaitTime = 60)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            if (element == null) return;
            var dropDownElement = new SelectElement(element);
            switch (selectBy.ToLower())
            {
                case "text": dropDownElement.SelectByText(strInputValue); break;
                case "index": dropDownElement.SelectByIndex(int.Parse(strInputValue)); break;
                case "value": dropDownElement.SelectByValue(strInputValue); break;
                default: dropDownElement.SelectByText(strInputValue); break;
            }
        }

        public List<string> GetDropdownListItems(By lookupBy, int maxWaitTime = 60)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            if (element == null) return new List<string>();
            var dropDownElement = new SelectElement(element);
            return dropDownElement.Options.Select(t => t.Text).ToList();
        }

        public List<int> GetDropdownListItemsByInt(By lookupBy, int maxWaitTime = 60)
        {
            Reporter.Add(new Act("Pick the time from application and match against time taken from SAT"));

            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            if (element == null) return new List<int>();
            var dropDownElement = new SelectElement(element);
            return dropDownElement.Options.Select(t => Convert.ToInt32(t.Text)).ToList();
        }

        public void ClearObjectValue(By lookupBy, int maxWaitTime = 60)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            element?.Clear();
        }

        public bool ValidateElementAttributeValue(By lookupBy, string attributeName, string expectedValue, int maxWaitTime = 60)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);

            return element != null && string.Equals(element.GetAttribute(attributeName).Trim().ToLower(),
                expectedValue.ToLower());
        }
        public bool IsElementSelected(By lookupBy, int maxWaitTime = 60)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);

            return element?.Selected ?? false;
        }

        public void MouseOverOnObject(By lookupBy, int maxWaitTime = 60)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            if (element == null) return;
            new Actions(Driver).MoveToElement(element).MoveByOffset(element.Size.Height / 2, element.Size.Height / 2).Perform();
            Thread.Sleep(1000);
        }

        public void DragAndDropObject(By lookupBy, By lookupBy1, int maxWaitTime = 60)
        {
            IWebElement element1 = null, element2 = null;

            element1 = WaitForElementVisible(lookupBy, maxWaitTime);
            element2 = WaitForElementVisible(lookupBy1, maxWaitTime);
            if (element1 == null || element2 == null) return;
            var builder = new Actions(Driver);

            var dragAndDrop = builder.ClickAndHold(element1)
                .MoveToElement(element2)
                .Release(element1)
                .Build();

            dragAndDrop.Perform();
            Thread.Sleep(1000);
        }

        public void MoveToElement(By lookupBy, int int_offSetX, int int_offSetY, int maxWaitTime = 60)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            if (element == null) return;
            new Actions(Driver).MoveToElement(element, int_offSetX, int_offSetY).Perform();
            Thread.Sleep(500);
        }
        public void MoveToElementAndClick(By lookupBy, int int_offSetX, int int_offSetY, int maxWaitTime = 60)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            if (element == null) return;
            MoveToElement(lookupBy, int_offSetX, int_offSetY);
            new Actions(Driver).Click().Perform();
            Thread.Sleep(500);
        }

        public bool CheckIfObjectExists(By lookupBy, int maxWaitTime = 30)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);

            return element != null;
        }

        public bool ValidateIfExists(By lookupBy, string objectName, int maxWaitTime = 30, bool condition = true)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);

            if (condition)
            {
                if (element != null)
                {
                    Reporter.Add(new Act($"Control '{objectName}' exist on the page"));
                    return true;
                }
                else
                {
                    Reporter.Add(new Act($"Control '{objectName}' doesn't exist on the page", false, Driver));
                    return false;
                }
            }
            else
            {
                if (element == null)
                {
                    Reporter.Add(new Act($"Control '{objectName}' doesn't exist on the page"));
                    return false;
                }
                else
                {
                    Reporter.Add(new Act($"Control '{objectName}' exist on the page", false, Driver));
                    return true;
                }
            }
        }

        /// <summary>
        /// Validate If Exists
        /// </summary>
        /// <param name="maxWaitTime"></param>
        /// <param name="condition"></param>
        /// <param name="objectName"></param>
        /// <param name="objects"></param>
        /// <returns></returns>
        public bool ValidateIfExists(string[] objectName, string[] objects, int maxWaitTime = 30, bool condition = true)
        {
            return objects.Select((t, i) => ValidateIfExists(By.XPath(t), objectName[i], maxWaitTime, condition))
                .All(x => x);
        }

        /// <summary>
        /// Validate control Value Or Text Or Enabled/Disabled
        /// </summary>
        /// <param name="lookupBy">Object to be verified</param>
        /// <param name="objectName">Name of the object</param>
        /// <param name="expectedTextOrValue">expected text/value. true for Enabled/Disabled state verification</param>
        /// <param name="attributeName">by default text. User can pass value, enabled or disabled</param>
        /// <param name="maxWaitTime">time to wait for the object visible</param>
        public void ValidateControlValueOrTextOrState(By lookupBy, string objectName, string expectedTextOrValue, string attributeName = "text", int maxWaitTime = 30)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            if (element != null)
            {
                if (attributeName.ToLower().Equals("text"))
                {
                    Reporter.Add(element.Text.Equals(expectedTextOrValue)
                        ? new Act(
                            $"control '{objectName}' text is shown as expected. The expected text is: '{expectedTextOrValue}'")
                        : new Act(
                            $"control '{objectName}' text is not shown as expected. The actual text is: '{element.Text}'", false, Driver));
                }
                else if (attributeName.ToLower().Equals("value"))
                {
                    Reporter.Add(element.GetAttribute("value").Equals(expectedTextOrValue)
                        ? new Act($"control '{objectName}' value is shown as expected")
                        : new Act(
                            $"control '{objectName}' value is not shown as expected. The actual value is: '{element.GetAttribute("value")}'", false, Driver));
                }
                else if (attributeName.ToLower().Equals("enabled"))
                {
                    Reporter.Add(element.Enabled.ToString().ToLower().Equals(expectedTextOrValue)
                        ? new Act($"control '{objectName}' is enabled as expected")
                        : new Act($"control '{objectName}' is disabled", false, Driver));
                }
                else if (attributeName.ToLower().Equals("disabled"))
                {
                    Reporter.Add(element.GetAttribute(attributeName).Equals(expectedTextOrValue)
                        ? new Act($"control {objectName} is disabled as expected")
                        : new Act($"control {objectName} is enabled", false, Driver));
                }
            }
            else
            {
                Reporter.Add(new Act($"Object '{objectName}' is not found to Verify TextOrValueOrState'", false, Driver));
            }
        }

        public void ValidateDropdownValue(By lookupBy, string objectName, string expectedTextOrValue, int maxWaitTime = 30)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);

            if (element == null) return;
            var dropDownElement = new SelectElement(element);
            Reporter.Add(dropDownElement.SelectedOption.Text.Equals(expectedTextOrValue)
                ? new Act(string.Format(
                    "control '{0}' value is shown as expected. The expected text is: " + expectedTextOrValue, objectName))
                : new Act($"control '{objectName}' value is not shown as expected. The actual value is: '{element.Text}'", false, Driver));
        }

        public void ObjectClick(By lookupBy, string objectName, int maxWaitTime = 60)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            if (element != null)
            {
                element.Click();
                Reporter.Add(new Act($"Clicked on '{objectName}'"));
            }

            WaitForPageLoad();
        }


        public void ObjectClickByPageScrolling(By lookupBy, string objectName, int YValue = 100, int maxWaitTime = 60)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            PageScrollDown(lookupBy, YValue);
            if (element == null) return;
            element.Click();
            Reporter.Add(new Act($"Clicked on '{objectName}'"));
        }

        public void ObjectClickByPageScrolling(IWebElement element, string objectName, int yValue = 100, int maxWaitTime = 60)
        {
            PageScrollDown(element, yValue);
            if (element == null) return;
            element.Click();
            Reporter.Add(new Act($"Clicked on '{objectName}'"));

        }

        /// <summary>
        /// Method to Set Object with Selected Input Value
        /// </summary>
        /// <param name="lookupBy">Object to loop for</param>
        /// <param name="objectName">Name of the object</param>
        /// <param name="inputValue">Input Value</param>
        /// <param name="maxWaitTime">Maximum Wait Time</param>
        public void SetObjectValue(By lookupBy, string objectName, string inputValue, int maxWaitTime = 60)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);

            if (element == null) return;
            element.Clear();
            element.SendKeys(inputValue);
            Reporter.Add(new Act(
                $"control '{objectName}' is set with value '{(objectName.ToLower().Equals("password") ? "*****" : inputValue)}'"));

        }

        public void DropdownSelectValue(By lookupBy, string objectName, string inputValue, string selectBy = "text", int maxWaitTime = 60)
        {
            IWebElement element = WaitForElementVisible(lookupBy, maxWaitTime);

            if (element != null)
            {
                SelectElement dropDownElement = new SelectElement(element);
                if (dropDownElement.Options.Count > 1)
                {
                    switch (selectBy.ToLower())
                    {
                        case "text": dropDownElement.SelectByText(inputValue); break;
                        case "index": dropDownElement.SelectByIndex(Convert.ToInt32(inputValue)); break;
                        case "value": dropDownElement.SelectByValue(inputValue); break;
                        default: dropDownElement.SelectByText(inputValue); break;
                    }

                    Reporter.Add(new Act(string.Format("control '{0}' is set with value '{1}'", objectName, inputValue)));
                }
            }

        }

        public string RetrieveObjectValue(By lookupBy, string objectName, string attributeName = "text")
        {
            const int maxWaitTime = 30;
            var element = WaitForElementVisible(lookupBy, maxWaitTime);

            if (element == null) return string.Empty;
            return attributeName.ToLower().Equals("text") ? element.Text : element.GetAttribute(attributeName);
        }

        public string RetrieveObjectValue(By lookupBy, string attributeName = "text")
        {
            const int maxWaitTime = 30;

            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            if (element == null) return string.Empty;
            return attributeName.ToLower().Equals("text") ? element.Text : element.GetAttribute(attributeName);
        }

        public string RetrieveDropdownValue(By lookupBy)
        {
            const int maxWaitTime = 30;

            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            return element == null ? string.Empty : new SelectElement(element).SelectedOption.Text;
        }

        /// <summary>
        /// Compare Lists
        /// </summary>
        /// <param name="actualList">ActualList</param>
        /// <param name="expectedList">ExpectedList</param>
        /// <returns>true, if validated successfully</returns>
        public bool CompareLists(List<string> actualList, List<string> expectedList, string nameOftheList = "")
        {
            var unmatchedList = actualList.Where(x => !expectedList.Contains(x)).ToList();

            if (unmatchedList.Count == 0)
            {
                Reporter.Add(string.IsNullOrEmpty(nameOftheList)
                    ? new Act("The Actual and Expected Lists are matched.")
                    : new Act($"The Actual and Expected List of {nameOftheList} are matched."));

                return true;
            }
            else
            {
                Reporter.Add(new Act("Following are the mismatches", false, Driver));
                unmatchedList.ForEach(x => Reporter.Add(new Act(x)));
                return false;
            }
        }

        /// <summary>
        /// Performs Click with JavaScript
        /// </summary>
        /// <param name="locator"></param>
        public void JavaScriptClick(By by, string objectName)
        {
            Reporter.Add(new Act($"Clicked on '{objectName}'"));
            var element = Driver.FindElement(by);
            var jsExecutor = (IJavaScriptExecutor)Driver;
            jsExecutor.ExecuteScript(@"arguments[0].click();", new object[] { element });
        }

        public bool ValidateDriverTitle(string strExpectedValue, int maxWaitTime = 60)
        {
            for (var i = 0; i < maxWaitTime; i++)
            {
                if (Driver.Title.ToString().Contains(strExpectedValue)) return true;
                else Thread.Sleep(1000);
            }

            return false;
        }

        public void NavigateToUrl(string strURL, string page = "")
        {
            Driver.Navigate().GoToUrl(strURL);
        }

        public string RetrieveCurrentBrowserUrl(int maxWaitTime = 0)
        {
            Thread.Sleep(maxWaitTime * 1000);

            return Driver.Url;
        }

        public void DoubleClickOnObject(By lookupBy, int maxWaitTime = 60)
        {
            IWebElement element = WaitForElementVisible(lookupBy, maxWaitTime);

            if (element == null) return;
            var action = new Actions(Driver);
            action.DoubleClick(element).Build().Perform();

        }

        public void SwitchToElement(By lookupBy, int maxWaitTime)
        {
            var element = WaitForElementVisible(lookupBy, maxWaitTime);
            if (element != null)
            {
                Driver.SwitchTo().Frame(element);
            }

        }

        public void SwitchToIFrames(By lookupBy)
        {
            try
            {
                IReadOnlyCollection<IWebElement> frames = Driver.FindElements(By.TagName("iframe"));
                foreach (var frame in frames)
                {
                    Driver.SwitchTo().Frame(frame);
                    IReadOnlyCollection<IWebElement> chframes = Driver.FindElements(By.TagName("iframe"));
                    if (chframes.Count.Equals(0))
                    {
                        Driver.SwitchTo().ParentFrame();
                    }
                    else
                    {
                        foreach (var chframe in chframes)
                        {
                            Driver.SwitchTo().Frame(chframe);
                            if (CheckIfObjectExists(lookupBy))
                            {
                                break;
                            }
                        }
                    }
                }
            }
            catch (StaleElementReferenceException ex)
            {
                Console.WriteLine(($"Failed at {MethodBase.GetCurrentMethod()} {ex.Message}\n{ex.StackTrace}"));
            }
        }

        public void SwitchToiFrame(By lookupBy, int maxWaitTime = 60)
        {
            var element = GetNativeElement(lookupBy, maxWaitTime);
            if (element == null) return;
            var webElement = GetNativeElementInElement(element, By.TagName("iframe"));
            Driver.SwitchTo().Frame(webElement);
        }

        public void SwitchToDefaultContent()
        {
            Driver.SwitchTo().DefaultContent();
        }

        public void SwitchToBaseWindow()
        {
            Driver.SwitchTo().Window(Driver.WindowHandles[0]);
        }

        /// <summary>
        /// Wait time
        /// </summary>
        /// <param name="maxWaitTime"></param>
        public void Wait(int maxWaitTime = 800)
        {
            Thread.Sleep(maxWaitTime);
        }

        public void WaitForPageLoad(int maxWaitTime = 800)
        {
            var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(maxWaitTime));
            var javascript = Driver as IJavaScriptExecutor;
            if (javascript == null)
                throw new ArgumentException("Driver", "Driver not supports javascript execution");
            wait.Until((d) =>
            {
                try
                {
                    string readyState = javascript.ExecuteScript(
                        "if (document.readyState) return document.readyState;").ToString();
                    return readyState.ToLower() == "complete";
                }
                catch (Exception ex)
                {
                    Reporter.Add(new Act(string.Format("Stopped execution after waiting for page to load for 10 sec"), false));
                    Console.WriteLine(
                        $"Stopped execution after waiting for page to load for 10 sec. Message {ex.Message}");
                    return false;
                }
            });
        }

        public void SwitchToWindow(string strWindowTitleName, string urlContent, int maxWaitTime = 60)
        {
            var isWindowFound = false;
            for (var i = 0; i < maxWaitTime; i++)
            {
                foreach (var window in Driver.WindowHandles.Where(x => x != Driver.CurrentWindowHandle))
                {
                    Driver.SwitchTo().Window(window);
                    if (Driver.Title != strWindowTitleName &&
                        !Driver.Url.ToLower().Contains(urlContent.ToLower())) continue;
                    isWindowFound = true;
                    break;
                }

                if (isWindowFound == true) break;
            }
            if (isWindowFound == false) throw new Exception(strWindowTitleName + " not found to switch");

        }

        public void SwitchToWindow(string windowTitleName, int maxWaitTime = 15)
        {
            Reporter.Add(new Act("Switching To Window"));
            Reporter.Add(new Act("Getting current window handle"));

            for (var i = 0; i < maxWaitTime; i++)
            {
                foreach (var window in Driver.WindowHandles.Where(x => x != Driver.CurrentWindowHandle).Select(y => y))
                {
                    Driver.SwitchTo().Window(window);
                    if (Driver.Title != windowTitleName) continue;
                    Reporter.Add(new Act(windowTitleName + " found"));
                    return;
                }

                Thread.Sleep(1000);
            }

            Reporter.Add(new Act(windowTitleName + " not found to switch"));
            throw new Exception(windowTitleName + " not found to switch");
        }

        public void SwitchToWindowContains(string strWindowTitleName, int maxWaitTime = 60)
        {
            for (var i = 0; i < maxWaitTime; i++)
            {
                var currentWindow = Driver.CurrentWindowHandle;

                foreach (var window in Driver.WindowHandles.Where(x => x != currentWindow))
                {
                    Driver.SwitchTo().Window(window);
                    if (Driver.Title.Contains(strWindowTitleName))
                    {
                        return;
                    }
                }

                Thread.Sleep(1000);
            }
            throw new Exception(strWindowTitleName + " not found to switch");
        }

        public void CloseWindow()
        {
            Driver.ExecuteScript("close()");
        }

        public void PageScrollTop()
        {
            var js = (IJavaScriptExecutor)Driver;
            js.ExecuteScript("window.scrollTo(0, 0)");
            Thread.Sleep(3000);
        }

        public void PageScrollUp(By by, int YValue = 100)
        {
            var element = WaitForElementVisible(by);

            if (element == null) return;
            IJavaScriptExecutor js = (IJavaScriptExecutor)Driver;
            js.ExecuteScript("arguments[0].scrollTop = arguments[1];", element, YValue);
            Thread.Sleep(3000);
        }

        public void PageScrollDown(By by, int YValue = 100)
        {
            IWebElement element = WaitForElementVisible(by);

            if (element == null) return;
            var js = (IJavaScriptExecutor)Driver;
            js.ExecuteScript($"window.scrollTo(0, {(element.Location.Y - YValue)})");
            Thread.Sleep(1000);
        }

        public void PageScrollDown(IWebElement element, int YValue = 100)
        {
            var js = (IJavaScriptExecutor)Driver;
            js.ExecuteScript($"window.scrollTo(0, {(element.Location.Y - YValue)})");
            Thread.Sleep(1000);
        }

        public string RandomString(int size)
        {
            var builder = new StringBuilder();
            var random = new Random();

            for (var i = 0; i < size; i++)
            {
                var ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));
                builder.Append(ch);
            }

            return builder.ToString();
        }

        public string GenerateRandomNumber(int size)
        {
            var random = new Random();
            var result = "";
            for (var i = 0; i < size; i++)
            {
                result += random.Next(0, 9).ToString();
            }
            return result;
        }

        /// <summary>
        /// Switch Between Windows & returns parent window handle
        /// </summary>
        public string SwitchToChildWindow()
        {
            var currentWindow = Driver.CurrentWindowHandle;

            Driver.SwitchTo().Window(Driver.WindowHandles.FirstOrDefault(x => x != currentWindow));

            return currentWindow;
        }

        /// <summary>
        /// Switch to window by taking window handle as parameter
        /// </summary>
        public void SwitchToWindowUsingWindowHandle(string windowHandle)
        {
            foreach (var handle in Driver.WindowHandles.Where(x => x.Equals(windowHandle)))
            {
                Driver.SwitchTo().Window(handle);
            }
        }

        public void LeaveCurrentWebSite()
        {
            Driver.SwitchTo().Alert().Accept();
        }

        /// <summary>
        /// Generates the date time string.
        /// </summary>
        /// <returns>Date Time stamp string</returns>
        public string GenerateDateTimeString()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }

        internal IAlert GetAlertHandle(int maxWaitTime = 10)
        {
            for (var i = 0; i < maxWaitTime; i++)
            {
                try
                {
                    return Driver.SwitchTo().Alert();
                }
                catch
                {
                    Thread.Sleep(1000);
                }
            }

            return null;
        }

        public bool IsAlertPresent()
        {
            var wait = new WebDriverWait(Driver, TimeSpan.FromMilliseconds(3000));
            return wait.Until(ExpectedConditions.AlertIsPresent()) != null;

        }

        public IWebElement WaitForElementExist(By by)
        {
            var wait = new WebDriverWait(Driver, TimeSpan.FromMilliseconds(3000));
            return wait.Until(ExpectedConditions.ElementExists(by));
        }

        public void AcceptAlert(int maxWaitTime = 10)
        {
            var alertHandle = GetAlertHandle(maxWaitTime);
            alertHandle?.Accept();
        }

        public void DismissAlert(int maxWaitTime = 10)
        {
            var alertHandle = GetAlertHandle(maxWaitTime);
            alertHandle?.Dismiss();
        }

        public string GetAlertText(int maxWaitTime = 10)
        {
            var alertHandle = GetAlertHandle(maxWaitTime);

            return alertHandle.Text ?? string.Empty;
        }

        public void CheckStringEqual(String actual, String expected)
        {
            Reporter.Add(new Act($"Verify '{expected}' Equals '{actual}'"));

            Reporter.Add(!string.Equals(actual, expected)
                         ? new Act($"Not Equal { actual } : {expected}", false, Driver)
                         : new Act($"{actual} : {expected} are equal"));
        }

        public void CheckStringContains(string expressionToSearch, string token)
        {
            Reporter.Add(new Act($"Verify '{expressionToSearch}' contains '{token}'"));
            if (!expressionToSearch.Contains(token))
            {
                Reporter.Add(new Act($"Does not Contain {expressionToSearch} : {token}", false, Driver));
            }
        }

        public void Equal(DateTime actual, DateTime expected, String name = "")
        {
            Reporter.Add(new Act($"Verify '{expected}' Equals '{actual}'"));

            if (!string.Equals(actual, expected))
            {
                throw new Exception($"Not Equal {actual} : {expected}");
            }
        }

        public void NullOrEmpty(String data)
        {
            Reporter.Add(new Act($"Verify Null or Empty '{data}'"));

            if (!string.IsNullOrEmpty(data) || !string.IsNullOrWhiteSpace(data))
            {
                throw new Exception(string.Format("Data is not Null or Empty"));
            }
        }

        public void NotNullOrEmpty(string data)
        {
            Reporter.Add(new Act($"Verify Null or Empty '{data}'"));

            if (string.IsNullOrEmpty(data) || string.IsNullOrWhiteSpace(data))
            {
                throw new Exception(string.Format("Data is Null or Empty"));
            }
        }

        public void Equal(long first, long second)
        {
            Reporter.Add(first.Equals(second)
                         ? new Act($"Verified '{first}' Equals '{second}'")
                         : new Act($"Not Equal {first} : {second}", false, Driver));

        }

        public void CheckEqualityOfObjects(bool first, bool second, string strStepDesc = "")
        {
            Reporter.Add(string.IsNullOrEmpty(strStepDesc)
                         ? new Act($"Verify '{first}' Equals '{second}'")
                         : new Act(strStepDesc));

            if (first != second)
            {
                throw new Exception($"Not Equal {first} : {second}");
            }
        }

        public void Equal(decimal first, decimal second, string strObject = "")
        {
            Reporter.Add(first != second
                ? new Act($"Expected {first} is not equal to Actual {second}", false, Driver)
                : new Act($"Expected '{first}' is equals to Actual '{second}'"));
        }

        public void GreaterThan(decimal first, decimal second)
        {
            if (first != 0)
            {
                Reporter.Add(first > second
                            ? new Act($"'{first}' Is Greater Than '{second}'")
                            : new Act($"'{first}' Is Less Than '{second}'", false, Driver));
            }
        }

        public void NotEqual(string first, string second)
        {
            Reporter.Add(new Act($"Verify '{first}' Not Equals '{second}'"));

            if (string.Equals(first, second))
            {
                throw new Exception($"Equal {first} : {second}");
            }
        }

        public void NotEqual(long first, long second)
        {
            Reporter.Add(new Act($"Verify '{first}' Not Equals '{second}'"));
        }

        public void PageReload()
        {
            Driver.Navigate().Refresh();
        }

        public void PageBack()
        {
            Driver.Navigate().Back();
        }

        public void PressEscapeKey()
        {
            Thread.Sleep(1000);
            var action = new Actions(Driver);
            action.SendKeys(OpenQA.Selenium.Keys.Escape);
            Thread.Sleep(1000);
        }

        public void PressArrowDownKey()
        {
            Thread.Sleep(1000);
            var action = new Actions(Driver);
            action.SendKeys(OpenQA.Selenium.Keys.ArrowDown).Build().Perform();
            Thread.Sleep(1000);
        }

        public void PressEnterKey()
        {
            Thread.Sleep(1000);
            var action = new Actions(Driver);
            action.SendKeys(OpenQA.Selenium.Keys.Enter).Build().Perform();
            Thread.Sleep(1000);
        }

        public void PressArrowUpKey()
        {
            Thread.Sleep(1000);
            var action = new Actions(Driver);
            action.SendKeys(OpenQA.Selenium.Keys.ArrowUp).Build().Perform();
            Thread.Sleep(1000);
        }

        public void PressTabKey()
        {
            Thread.Sleep(1000);
            var action = new Actions(Driver);
            action.SendKeys(OpenQA.Selenium.Keys.Tab).Build().Perform();
            Thread.Sleep(1000);
        }
    }
}