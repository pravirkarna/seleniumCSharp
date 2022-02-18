using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Automation.Ui.Accelerators.ReportingClassess
{
    public class Summary
    {
        public List<TestCase> TestCases { get; } = new List<TestCase>();

        public int PassedCount => TestCases.FindAll(i => i.IsSuccess == true).Count;

        public int FailedCount=> TestCases.FindAll(i => i.IsSuccess == false).Count;

        public bool IsSuccess=> TestCases.FindAll(i => i.IsSuccess == false).Count == 0;
        
        public Dictionary<String, Dictionary<String, long>> GetStatusByBrowser()
        {
            var result = new Dictionary<string, Dictionary<string, long>>();

            foreach (TestCase testCase in this.TestCases)
            {
                foreach (Browser browser in testCase.Browsers)
                {
                    var browserQaulifier = $"{browser.BrowserName} {browser.BrowserVersion}";
                    if (!result.Keys.Contains(browserQaulifier))
                    {
                        result.Add(browserQaulifier, new Dictionary<string, long>());
                        result[browserQaulifier].Add("PASSED", 0);
                        result[browserQaulifier].Add("FAILED", 0);
                    }

                    result[browserQaulifier]["PASSED"] += browser.PassedCount;
                    result[browserQaulifier]["FAILED"] += browser.FailedCount;
                }
            }

            return result;
        }
    }
}