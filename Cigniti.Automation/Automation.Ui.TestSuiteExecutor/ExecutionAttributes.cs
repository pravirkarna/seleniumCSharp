using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Maxim.Automation.Ui.TestSuiteExecutor
{
    public class ExecutionAttributes
    {
        public string ExecutionCategories { get; set; } = string.Empty;
       
        public string TestCaseDescription { get; set; } = string.Empty;
       
        public string TestCaseName { get; set; } = string.Empty;
       
        public string TCID { get; set; } = string.Empty;
        
        public string UserStoryId { get; set; } = string.Empty;
        
        public string SubModuleName { get; set; } = string.Empty;
        
        public string ModuleName { get; set; } = string.Empty;
    }
}
