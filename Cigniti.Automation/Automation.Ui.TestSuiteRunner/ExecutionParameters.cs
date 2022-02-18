using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Maxim.Automation.Ui.TestSuiteRunner
{
   public class ExecutionParameters
    {
        public string ApplicationName { get; set; } = "ETrak";
        public string SubModule { get; set; } = "All";
        public string TestCaseNames { get; set; } = "All";
        public string ExecutionCatogery { get; set; } = "regression";
        public string Environment { get; set; } = "RC";
        public string MaxParallel { get; set; } = "1";
        public string ReportingPath { get; set; } = @"C:\Reports";
        public string Browser { get; set; } = "chrome";
    }
}
