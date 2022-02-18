using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Automation.Ui.Accelerators.UtilityClasses
{
    public static class SeleniumExtension
    {
        /// <summary>
        /// HighLights elements in web page.
        /// </summary>
        /// <param name="element"></param>
        /// <param name="driver"></param>
        public static IWebElement HighLightElement(this IWebElement element, RemoteWebDriver driver, string color="orange")
        {

            if (element == null)
                return null;

            var highlightJs =
                $@"arguments[0].style.cssText = ""border-width: 4px; border-style: solid; border-color: {color}"";";
            var elementToHighlight = new object[] { element };
            var scrollToViewScript = "$(arguments[0].scrollIntoView(true));";

            try
            {
                var jsExecutor = (IJavaScriptExecutor)driver;
                jsExecutor.ExecuteScript(highlightJs, elementToHighlight);
                jsExecutor.ExecuteScript(scrollToViewScript, elementToHighlight);
            }
            catch
            {
                // ignored
            }

            return element;
        }
    }
}