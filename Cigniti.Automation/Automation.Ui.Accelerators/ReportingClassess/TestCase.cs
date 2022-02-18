using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Automation.Ui.Accelerators.ReportingClassess
{
    public class TestCase
    {
        public TestCase(string moduleName, string subModuleName, string userStoryId, string testCaseId, string id, string name, string requirementFeature, string executionCategory)
        {
            this.ModuleName = moduleName;
            this.SubModuleName = subModuleName;
            this.UserStory = userStoryId;
            this.TestCaseId = testCaseId;
            this.Title = id;
            this.Name = name;
            this.RequirementFeature = requirementFeature;
            this.ExecutionCategory = executionCategory;
        }

        public TestCase(TestCase testCase)
        {
            ModuleName = testCase.ModuleName;
            SubModuleName = testCase.SubModuleName;
            UserStory = testCase.UserStory;
            TestCaseId = testCase.TestCaseId;
            Title = testCase.Name;
            Name = testCase.Name;
            RequirementFeature = testCase.RequirementFeature;
            ExecutionCategory = testCase.ExecutionCategory;
        }

        public TestCase()
        {

        }

        /// <summary>
        /// Gets or sets Module
        /// </summary>
        public string ModuleName { get; set; }


        /// <summary>
        /// Gets or sets SubModuleName
        /// </summary>
        public string SubModuleName { get; set; }

        /// <summary>
        /// Gets or sets Title
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets Test Case ID
        /// </summary>
        public string TestCaseId { get; set; }

        /// <summary>
        /// Gets or sets User Story ID
        /// </summary>
        public string UserStory { get; set; }

        /// <summary>
        /// Gets or sets Test Case Name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets Requirement Feature
        /// </summary>
        public string RequirementFeature { get; set; }

        /// <summary>
        /// Gets or sets Test Case Execution Category
        /// </summary>
        public string ExecutionCategory { get; set; }

        /// <summary>
        /// Gets Browsers
        /// </summary>
        public List<Browser> Browsers { get; } = new List<Browser>();

        /// <summary>
        /// Gets current Browser
        /// </summary>
        public Browser Browser => Browsers.Last();

        /// <summary>
        /// Gets Passed Count
        /// </summary>
        public int PassedCount => Browsers.Count(i =>i.IsSuccess.Equals(true));

        /// <summary>
        /// Gets Failed Count
        /// </summary>
        public int FailedCount => Browsers.Count(i => i.IsSuccess.Equals(false));

        /// <summary>
        /// Gets IsSuccess
        /// </summary>
        public bool IsSuccess=> Browsers.Any(i => i.IsSuccess == true);
        
        /// <summary>
        /// Gets or sets BugInfo
        /// </summary>
        public string BugInfo { get; set; }

        public Summary Summary { get; set; }

        public string BrowserName { get; set; }
    }
}